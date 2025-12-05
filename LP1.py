import sys
import os
import json
import time
import threading
import re
import datetime  # 用於生成日誌檔名的日期
import winsound  # Windows 音效

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
#   爬蟲執行緒模組 (Crawler Threads)
# ==========================================

class BaseCrawlerThread(QThread):
    log_signal = pyqtSignal(str)
    price_signal = pyqtSignal(str, float, float, str)  # (Source, Bid, Ask, Time)
    status_signal = pyqtSignal(str, str)  # (Source, Status Msg)
    finished_signal = pyqtSignal(str)  # (Source)

    def __init__(self, source_name):
        super().__init__()
        self.source_name = source_name
        self.running = True
        self.driver = None

    def setup_driver(self):
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--log-level=3")
        chrome_options.add_argument(
            "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

        self.driver = webdriver.Chrome(options=chrome_options)

    def stop(self):
        self.running = False

    def stop_driver(self):
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
            self.driver = None


# --- 1. 永豐金爬蟲 (Wing Fung) ---
class WFThread(BaseCrawlerThread):
    def run(self):
        try:
            self.status_signal.emit(self.source_name, "啟動中...")
            self.setup_driver()
            self.log_signal.emit(f"[{self.source_name}] Driver 就緒，前往網站...")

            self.driver.get("https://www.wfbullion.com/")
            wait = WebDriverWait(self.driver, 20)

            while self.running:
                try:
                    price_element = wait.until(EC.presence_of_element_located((By.ID, "pm-llg")))
                    raw_text = price_element.text.strip()
                    lines = raw_text.split('\n')

                    if len(lines) > 3:
                        bid_str = lines[2].strip().replace(',', '')
                        ask_line = lines[3].strip()
                        ask_str = ask_line.split(' ')[0].replace(',', '')

                        bid = float(bid_str)
                        ask = float(ask_str)
                        now_str = time.strftime("%H:%M:%S")

                        self.price_signal.emit(self.source_name, bid, ask, now_str)
                        self.status_signal.emit(self.source_name, "監控中")
                    else:
                        self.status_signal.emit(self.source_name, "數據格式異常")

                except Exception as e:
                    time.sleep(2)

                for _ in range(10):
                    if not self.running: break
                    time.sleep(0.1)

        except Exception as e:
            self.log_signal.emit(f"[{self.source_name}] 錯誤: {str(e)}")
            self.status_signal.emit(self.source_name, "停止")
        finally:
            self.stop_driver()
            self.finished_signal.emit(self.source_name)


# --- 2. IG Markets 爬蟲 ---
class IGThread(BaseCrawlerThread):
    def run(self):
        try:
            self.status_signal.emit(self.source_name, "啟動中...")
            self.setup_driver()
            self.driver.get("https://www.ig.com/cn/commodities/markets-commodities/gold")
            wait = WebDriverWait(self.driver, 20)

            while self.running:
                try:
                    sell_el = wait.until(EC.presence_of_element_located(
                        (By.CSS_SELECTOR, ".price-ticket__button--sell .price-ticket__price")))
                    buy_el = wait.until(EC.presence_of_element_located(
                        (By.CSS_SELECTOR, ".price-ticket__button--buy .price-ticket__price")))

                    bid = parse_price(sell_el.text)
                    ask = parse_price(buy_el.text)
                    now_str = time.strftime("%H:%M:%S")

                    if bid > 0 and ask > 0:
                        self.price_signal.emit(self.source_name, bid, ask, now_str)
                        self.status_signal.emit(self.source_name, "監控中")

                except Exception as e:
                    self.status_signal.emit(self.source_name, "等待元素...")
                    time.sleep(2)

                for _ in range(10):
                    if not self.running: break
                    time.sleep(0.1)

        except Exception as e:
            self.log_signal.emit(f"[{self.source_name}] 錯誤: {str(e)}")
        finally:
            self.stop_driver()
            self.finished_signal.emit(self.source_name)


# --- 3. Oanda 爬蟲 ---
class OandaThread(BaseCrawlerThread):
    def run(self):
        try:
            self.status_signal.emit(self.source_name, "啟動中...")
            self.setup_driver()
            self.driver.get("https://www.oanda.com/bvi-en/cfds/metals/")
            wait = WebDriverWait(self.driver, 20)

            while self.running:
                try:
                    gold_row = wait.until(
                        EC.presence_of_element_located((By.XPATH, "//tr[.//span[contains(text(), 'Gold')]]")))
                    cells = gold_row.find_elements(By.TAG_NAME, "td")

                    if len(cells) >= 3:
                        bid = parse_price(cells[1].text)
                        ask = parse_price(cells[2].text)
                        now_str = time.strftime("%H:%M:%S")

                        if bid > 0 and ask > 0:
                            self.price_signal.emit(self.source_name, bid, ask, now_str)
                            self.status_signal.emit(self.source_name, "監控中")

                except Exception as e:
                    self.status_signal.emit(self.source_name, "重試中...")
                    time.sleep(2)

                for _ in range(10):
                    if not self.running: break
                    time.sleep(0.1)

        except Exception as e:
            self.log_signal.emit(f"[{self.source_name}] 錯誤: {str(e)}")
        finally:
            self.stop_driver()
            self.finished_signal.emit(self.source_name)


# --- 4. Forex.com 爬蟲 ---
class ForexThread(BaseCrawlerThread):
    def run(self):
        try:
            self.status_signal.emit(self.source_name, "啟動中...")
            self.setup_driver()
            self.driver.get("https://www.forex.com/cn/markets-to-trade/precious-metals/")
            wait = WebDriverWait(self.driver, 20)

            while self.running:
                try:
                    product_row = wait.until(EC.presence_of_element_located((By.XPATH, "//tr[.//a[@title='XAU USD']]")))
                    bid_element = product_row.find_element(By.CSS_SELECTOR, ".mp__td--Bid")
                    offer_element = product_row.find_element(By.CSS_SELECTOR, ".mp__td--Offer")

                    bid = parse_price(bid_element.text)
                    ask = parse_price(offer_element.text)
                    now_str = time.strftime("%H:%M:%S")

                    if bid > 0 and ask > 0:
                        self.price_signal.emit(self.source_name, bid, ask, now_str)
                        self.status_signal.emit(self.source_name, "監控中")

                except Exception as e:
                    self.status_signal.emit(self.source_name, "重試中...")
                    time.sleep(2)

                for _ in range(10):
                    if not self.running: break
                    time.sleep(0.1)

        except Exception as e:
            self.log_signal.emit(f"[{self.source_name}] 錯誤: {str(e)}")
        finally:
            self.stop_driver()
            self.finished_signal.emit(self.source_name)


# ==========================================
#   UI 元件
# ==========================================

class BrokerPanel(QGroupBox):
    """顯示單一券商報價的面板"""

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
        self.setWindowTitle("全方位黃金監控系統 (日誌存檔版)")
        self.resize(1000, 800)

        self.threads = {}
        self.setting_inputs = {}
        self.alert_status_labels = {}

        # 記錄上一次觸發的層級 (-1:無, 0:層級1, 1:層級2, 2:層級3)
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

        # 頂部控制列
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

        # 主要分頁區
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        # 分頁 1: 監控儀表板 (2x2)
        self.tab_monitor = QWidget()
        self.setup_monitor_tab()
        self.tabs.addTab(self.tab_monitor, "監控儀表板")

        # 分頁 2: 獨立警報設定
        self.tab_settings = QWidget()
        self.setup_settings_tab()
        self.tabs.addTab(self.tab_settings, "各平台警報設定")

        # 分頁 3: 運行日誌
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
        layout.addWidget(QLabel("請在此為每一間公司分別設定 3 層級的點差警報。設定後請點擊「儲存設定」或直接開始監控。"))
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
        layout.addWidget(QLabel("音效路徑 (建議使用 .wav)"), 0, 2)
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

            self.setting_inputs[key].append({
                "diff": txt_diff,
                "sound": txt_sound
            })
        layout.setRowStretch(4, 1)
        parent_tab.addTab(page, title)

    def setup_log_tab(self):
        layout = QVBoxLayout(self.tab_log)
        self.txt_log = QTextEdit()
        self.txt_log.setReadOnly(True)
        layout.addWidget(self.txt_log)
        btn_clear = QPushButton("清除介面日誌 (不影響檔案)")
        btn_clear.clicked.connect(self.txt_log.clear)
        layout.addWidget(btn_clear)

    def update_realtime_clock(self):
        self.lbl_clock.setText(QTime.currentTime().toString("HH:mm:ss"))

    def browse_file(self, line_edit):
        f, _ = QFileDialog.getOpenFileName(self, "選取音效", "", "WAV Audio (*.wav);;All (*.*)")
        if f: line_edit.setText(f)

    @pyqtSlot(str)
    def log_message(self, msg):
        """
        記錄日誌功能：
        1. 顯示在 UI 的 TextEdit
        2. 寫入到本地的 txt 檔案
        """
        # 取得當前時間 HH:MM:SS
        ts = time.strftime("%H:%M:%S")
        full_log_text = f"[{ts}] {msg}"

        # 1. 更新 UI
        self.txt_log.append(full_log_text)

        # 2. 寫入檔案 (檔名格式: monitor_log_2023-10-27.txt)
        today_date = datetime.datetime.now().strftime("%Y-%m-%d")
        log_filename = f"monitor_log_{today_date}.txt"

        try:
            with open(log_filename, "a", encoding="utf-8") as f:
                f.write(full_log_text + "\n")
        except Exception as e:
            # 如果寫入檔案失敗，至少在控制台印出
            print(f"寫入日誌檔案失敗: {e}")

    def start_monitor(self):
        self.save_settings()
        self.btn_start.setEnabled(False)
        self.btn_stop.setEnabled(True)

        # 重置所有觸發狀態
        self.last_triggered_levels = {}

        self.log_message("--- 系統啟動 (使用自動 Driver 管理) ---")

        self.start_thread("WF", WFThread)
        self.start_thread("IG", IGThread)
        self.start_thread("Oanda", OandaThread)
        self.start_thread("Forex", ForexThread)

    def start_thread(self, key, thread_class):
        t = thread_class(key)
        t.log_signal.connect(self.log_message)
        t.price_signal.connect(self.on_price_update)
        t.status_signal.connect(self.on_status_update)
        t.finished_signal.connect(self.on_thread_finished)

        self.threads[key] = t
        t.start()

    def stop_monitor(self):
        self.log_message("正在停止所有監控...")
        self.btn_stop.setEnabled(False)
        for t in self.threads.values():
            t.stop()

    def on_price_update(self, source, bid, ask, time_str):
        spread = 0.0
        if source == "WF":
            spread = self.panel_wf.update_data(bid, ask, time_str)
        elif source == "IG":
            spread = self.panel_ig.update_data(bid, ask, time_str)
        elif source == "Oanda":
            spread = self.panel_oanda.update_data(bid, ask, time_str)
        elif source == "Forex":
            spread = self.panel_forex.update_data(bid, ask, time_str)

        self.check_alert(source, spread)

    def on_status_update(self, source, msg):
        if source == "WF":
            self.panel_wf.update_status(msg)
        elif source == "IG":
            self.panel_ig.update_status(msg)
        elif source == "Oanda":
            self.panel_oanda.update_status(msg)
        elif source == "Forex":
            self.panel_forex.update_status(msg)

    def on_thread_finished(self, source):
        self.log_message(f"[{source}] 停止運作")
        if not any(t.isRunning() for t in self.threads.values()):
            self.btn_start.setEnabled(True)
            self.btn_stop.setEnabled(False)
            self.log_message("--- 所有監控已結束 ---")

    def check_alert(self, source, spread):
        inputs = self.setting_inputs.get(source, [])

        # 初始化：當前最高觸發的層級索引 (-1 代表無)
        current_highest_level = -1
        current_sound_path = None
        current_threshold_str = "0.00"

        # 1. 掃描所有層級，找出目前點差符合的最高層級
        for i, item in enumerate(inputs):
            try:
                threshold_text = item['diff'].text()
                threshold = float(threshold_text) if threshold_text else 0.0
            except:
                threshold = 0.0

            status_lbl = self.alert_status_labels.get((source, i))

            if threshold > 0 and spread >= threshold:
                status_lbl.setText("觸發中")
                status_lbl.setStyleSheet("color: red; font-weight: bold;")

                current_highest_level = i
                current_sound_path = item['sound'].text()
                current_threshold_str = threshold_text  # 保存觸發的設定值
            else:
                status_lbl.setText("正常")
                status_lbl.setStyleSheet("color: green;")

        # 2. 獲取上一次記錄的狀態
        last_level = self.last_triggered_levels.get(source, -1)

        # 3. 核心邏輯判斷 (升級觸發)
        if current_highest_level > last_level:

            # --- 格式化日誌訊息 (嚴格依照要求) ---
            # 格式: [Source] 點差:當前點差 大於 層級點差:設定值
            formatted_msg = f"[{source}] 點差:{spread:.2f} 大於 層級點差:{current_threshold_str}"
            self.log_message(formatted_msg)

            # 播放聲音
            if current_sound_path:
                threading.Thread(target=self.play_sound, args=(current_sound_path,), daemon=True).start()

            # 更新記憶狀態
            self.last_triggered_levels[source] = current_highest_level

        elif current_highest_level < last_level:
            # 降級了，更新狀態但不播放聲音
            self.last_triggered_levels[source] = current_highest_level

    def play_sound(self, path):
        normalized_path = os.path.normpath(path)
        if os.path.exists(normalized_path):
            try:
                winsound.PlaySound(normalized_path, winsound.SND_FILENAME | winsound.SND_NODEFAULT)
            except Exception as e:
                self.audio_log_signal.emit(f"音效播放失敗: {str(e)}")
        else:
            self.audio_log_signal.emit(f"找不到音效檔案: {normalized_path}")

    def save_settings(self):
        data = {}
        for key, inputs in self.setting_inputs.items():
            data[key] = []
            for item in inputs:
                data[key].append({
                    "diff": item['diff'].text(),
                    "sound": item['sound'].text()
                })
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            self.log_message("設定已儲存")
        except Exception as e:
            self.log_message(f"儲存失敗: {e}")

    def load_settings(self):
        if not os.path.exists(CONFIG_FILE): return
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            for key, tiers in data.items():
                if key in self.setting_inputs:
                    ui_inputs = self.setting_inputs[key]
                    for i, t_data in enumerate(tiers):
                        if i < len(ui_inputs):
                            ui_inputs[i]['diff'].setText(t_data.get('diff', ''))
                            ui_inputs[i]['sound'].setText(t_data.get('sound', ''))
        except Exception as e:
            self.log_message(f"讀取設定失敗: {e}")

    def closeEvent(self, event):
        if any(t.isRunning() for t in self.threads.values()):
            reply = QMessageBox.question(self, '確認', '程式正在執行，確定要關閉嗎？',
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Yes:
                self.stop_monitor()
                event.accept()
            else:
                event.ignore()
        else:
            self.save_settings()
            event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setFont(QFont("Microsoft JhengHei", 10))
    window = GoldMonitorApp()
    window.show()
    sys.exit(app.exec())