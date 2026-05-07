import sys
import os
import json
import csv
import time
import requests
import subprocess
from datetime import datetime
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                               QTableWidget, QTableWidgetItem, QHeaderView, QDoubleSpinBox,
                               QMessageBox, QTabWidget, QComboBox, QSizePolicy, QFileDialog)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QIcon, QColor

# --- 版本資訊 ---
CURRENT_VERSION = "1.0.0"
UPDATE_URL = "https://raw.githubusercontent.com/jung71219-sys/game_logger/refs/heads/main/game_data.json"

def resource_path(relative_path):
    """ 取得資源絕對路徑，支援 PyInstaller 打包環境 """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

class GameTracker(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"遊戲練功效率紀錄工具 v{CURRENT_VERSION}")
        self.resize(1200, 900)
        
        self.data_file = "game_data.json"
        self.start_time = None
        self.timer_display = QTimer()
        self.timer_display.timeout.connect(self.sync_realtime_updates)

        self.equip_data = {
            "武器": [], "項鍊": [], "卡片": [], 
            "坐騎": [], "鬥魂": [], "寵物": []
        }

        icon_file = resource_path("Exp.ico")
        if os.path.exists(icon_file):
            self.setWindowIcon(QIcon(icon_file))

        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        self.record_tab = QWidget()
        self.analysis_tab = QWidget()
        self.config_tab = QWidget()

        self.tabs.addTab(self.record_tab, "練功紀錄")
        self.tabs.addTab(self.analysis_tab, "數據分析")
        self.tabs.addTab(self.config_tab, "裝備設定")

        self.setup_record_tab()
        self.setup_analysis_tab()
        self.setup_config_tab()
        
        self.load_data()
        self.is_dark_mode = False
        self.toggle_dark_mode()
        
        QTimer.singleShot(3000, self.auto_check_update)

    def setup_record_tab(self):
        layout = QVBoxLayout(self.record_tab)
        layout.setSpacing(10)
        plus_levels = [f"+{i}" for i in range(13)]

        def set_resizable(widget):
            widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        tool_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("搜尋備註或裝備名稱...")
        self.search_input.textChanged.connect(self.filter_table)
        
        self.dark_mode_btn = QPushButton("切換深淺模式")
        self.dark_mode_btn.clicked.connect(self.toggle_dark_mode)
        
        self.import_csv_btn = QPushButton("匯入 CSV")
        self.import_csv_btn.clicked.connect(self.import_from_csv)
        
        self.export_csv_btn = QPushButton("匯出 CSV")
        self.export_csv_btn.clicked.connect(self.export_to_csv)
        
        tool_layout.addWidget(QLabel("搜尋:"))
        tool_layout.addWidget(self.search_input)
        tool_layout.addWidget(self.dark_mode_btn)
        tool_layout.addWidget(self.import_csv_btn)
        tool_layout.addWidget(self.export_csv_btn)

        input_layout1 = QHBoxLayout()
        self.weapon_plus = QComboBox(); self.weapon_plus.addItems(plus_levels); self.weapon_plus.setFixedWidth(55)
        self.weapon_input = QComboBox(); set_resizable(self.weapon_input)
        self.neck_plus = QComboBox(); self.neck_plus.addItems(plus_levels); self.neck_plus.setFixedWidth(55)
        self.neck_input = QComboBox(); set_resizable(self.neck_input)
        self.mount_plus = QComboBox(); self.mount_plus.addItems(plus_levels); self.mount_plus.setFixedWidth(55)
        self.mount_input = QComboBox(); set_resizable(self.mount_input)
        self.soul_plus = QComboBox(); self.soul_plus.addItems(plus_levels); self.soul_plus.setFixedWidth(55)
        self.soul_input = QComboBox(); set_resizable(self.soul_input)

        input_layout1.addWidget(QLabel("武器:")); input_layout1.addWidget(self.weapon_plus); input_layout1.addWidget(self.weapon_input)
        input_layout1.addWidget(QLabel("項鍊:")); input_layout1.addWidget(self.neck_plus); input_layout1.addWidget(self.neck_input)
        input_layout1.addWidget(QLabel("坐騎:")); input_layout1.addWidget(self.mount_plus); input_layout1.addWidget(self.mount_input)
        input_layout1.addWidget(QLabel("鬥魂:")); input_layout1.addWidget(self.soul_plus); input_layout1.addWidget(self.soul_input)

        input_layout_cards = QHBoxLayout()
        self.card_plus1 = QComboBox(); self.card_plus1.addItems(plus_levels); self.card_plus1.setFixedWidth(55)
        self.card_input1 = QComboBox(); set_resizable(self.card_input1)
        self.card_plus2 = QComboBox(); self.card_plus2.addItems(plus_levels); self.card_plus2.setFixedWidth(55)
        self.card_input2 = QComboBox(); set_resizable(self.card_input2)
        self.card_plus3 = QComboBox(); self.card_plus3.addItems(plus_levels); self.card_plus3.setFixedWidth(55)
        self.card_input3 = QComboBox(); set_resizable(self.card_input3)
        self.card_plus4 = QComboBox(); self.card_plus4.addItems(plus_levels); self.card_plus4.setFixedWidth(55)
        self.card_input4 = QComboBox(); set_resizable(self.card_input4)
        self.pet_input = QComboBox(); set_resizable(self.pet_input)

        combos = [self.weapon_input, self.neck_input, self.card_input1, self.card_input2,
                  self.card_input3, self.card_input4, self.mount_input, self.soul_input, self.pet_input]
        for combo in combos:
            combo.setEditable(True)

        input_layout_cards.addWidget(QLabel("卡1:")); input_layout_cards.addWidget(self.card_plus1); input_layout_cards.addWidget(self.card_input1)
        input_layout_cards.addWidget(QLabel("卡2:")); input_layout_cards.addWidget(self.card_plus2); input_layout_cards.addWidget(self.card_input2)
        input_layout_cards.addWidget(QLabel("卡3:")); input_layout_cards.addWidget(self.card_plus3); input_layout_cards.addWidget(self.card_input3)
        input_layout_cards.addWidget(QLabel("卡4:")); input_layout_cards.addWidget(self.card_plus4); input_layout_cards.addWidget(self.card_input4)
        input_layout_cards.addWidget(QLabel("寵物:")); input_layout_cards.addWidget(self.pet_input)

        input_layout2 = QHBoxLayout()
        self.crit_input = QLineEdit(); set_resizable(self.crit_input); self.crit_input.setPlaceholderText("爆傷 %")
        self.atk_boost_input = QLineEdit(); set_resizable(self.atk_boost_input); self.atk_boost_input.setPlaceholderText("攻擊力增幅")
        self.time_input = QDoubleSpinBox(); set_resizable(self.time_input); self.time_input.setRange(0, 999999); self.time_input.setSuffix(" 分鐘")
        self.exp_start_input = QDoubleSpinBox(); set_resizable(self.exp_start_input); self.exp_start_input.setRange(0, 100); self.exp_start_input.setDecimals(5); self.exp_start_input.setSuffix(" %")
        self.exp_end_input = QDoubleSpinBox(); set_resizable(self.exp_end_input); self.exp_end_input.setRange(0, 100); self.exp_end_input.setDecimals(5); self.exp_end_input.setSuffix(" %")
        self.note_input = QLineEdit(); set_resizable(self.note_input); self.note_input.setPlaceholderText("備註")

        input_layout2.addWidget(QLabel("爆傷:")); input_layout2.addWidget(self.crit_input)
        input_layout2.addWidget(QLabel("攻增:")); input_layout2.addWidget(self.atk_boost_input)
        input_layout2.addWidget(QLabel("時長:")); input_layout2.addWidget(self.time_input)
        input_layout2.addWidget(QLabel("起始:")); input_layout2.addWidget(self.exp_start_input)
        input_layout2.addWidget(QLabel("結束:")); input_layout2.addWidget(self.exp_end_input)
        input_layout2.addWidget(QLabel("備註:")); input_layout2.addWidget(self.note_input)

        btn_layout = QHBoxLayout()
        self.timer_btn = QPushButton("開始計時")
        self.timer_btn.setStyleSheet("background-color: #673AB7; color: white; font-weight: bold; height: 35px;")
        self.timer_btn.clicked.connect(self.toggle_timer)
        self.copy_btn = QPushButton("複製前一筆"); self.copy_btn.clicked.connect(self.copy_last_record)
        self.clear_btn = QPushButton("一鍵清空"); self.clear_btn.clicked.connect(self.clear_inputs)
        add_btn = QPushButton("新增紀錄"); add_btn.setStyleSheet("background-color: #4CAF50; color: white;"); add_btn.clicked.connect(self.add_record)
        edit_btn = QPushButton("修改選中"); edit_btn.setStyleSheet("background-color: #2196F3; color: white;"); edit_btn.clicked.connect(self.edit_record)
        del_btn = QPushButton("刪除選中"); del_btn.setStyleSheet("background-color: #f44336; color: white;"); del_btn.clicked.connect(self.delete_record)

        btn_layout.addWidget(self.timer_btn); btn_layout.addWidget(self.copy_btn); btn_layout.addWidget(self.clear_btn)
        btn_layout.addWidget(add_btn); btn_layout.addWidget(edit_btn); btn_layout.addWidget(del_btn)

        self.table = QTableWidget()
        self.table.setColumnCount(17)
        self.table.setHorizontalHeaderLabels(["武器", "項鍊", "卡1", "卡2", "卡3", "卡4", "坐騎", "鬥魂", "寵物", "爆傷", "攻增", "時間", "前%", "後%", "獲得%", "時薪%", "備註"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.itemDoubleClicked.connect(self.load_to_inputs)

        layout.addLayout(tool_layout); layout.addLayout(input_layout1); layout.addLayout(input_layout_cards); layout.addLayout(input_layout2); layout.addLayout(btn_layout); layout.addWidget(self.table)

    def setup_analysis_tab(self):
        main_layout = QVBoxLayout(self.analysis_tab)
        countdown_group = QHBoxLayout()
        self.next_lvl_exp = QDoubleSpinBox()
        self.next_lvl_exp.setRange(0, 100); self.next_lvl_exp.setDecimals(5); self.next_lvl_exp.setPrefix("尚差: "); self.next_lvl_exp.setSuffix(" %")
        self.next_lvl_exp.valueChanged.connect(self.calculate_countdown)
        self.countdown_label = QLabel("預計升級所需時間: --"); self.countdown_label.setStyleSheet("font-weight: bold; color: #FF9800;")
        countdown_group.addWidget(self.next_lvl_exp); countdown_group.addWidget(self.countdown_label)
        main_layout.addLayout(countdown_group)
        self.rank_table = QTableWidget(); self.rank_table.setColumnCount(5)
        self.rank_table.setHorizontalHeaderLabels(["排名", "組合", "平均爆傷", "最高時薪", "次數"])
        self.rank_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        main_layout.addWidget(self.rank_table)
        refresh_rank_btn = QPushButton("刷新分析"); refresh_rank_btn.clicked.connect(self.update_analysis)
        main_layout.addWidget(refresh_rank_btn)

    def setup_config_tab(self):
        layout = QVBoxLayout(self.config_tab)
        config_search_layout = QHBoxLayout()
        self.config_search_input = QLineEdit(); self.config_search_input.setPlaceholderText("搜尋...")
        self.config_search_input.textChanged.connect(self.filter_config_table)
        config_search_layout.addWidget(QLabel("搜尋:")); config_search_layout.addWidget(self.config_search_input)
        
        config_input_layout = QHBoxLayout()
        self.cate_combo = QComboBox(); self.cate_combo.addItems(["武器", "項鍊", "卡片", "坐騎", "鬥魂", "寵物"])
        self.item_name_input = QLineEdit(); add_item_btn = QPushButton("新增")
        add_item_btn.clicked.connect(self.add_config_item)
        config_input_layout.addWidget(QLabel("分類:")); config_input_layout.addWidget(self.cate_combo)
        config_input_layout.addWidget(self.item_name_input); config_input_layout.addWidget(add_item_btn)

        self.config_table = QTableWidget(); self.config_table.setColumnCount(2)
        self.config_table.setHorizontalHeaderLabels(["類別", "名稱"]); self.config_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.config_table.itemChanged.connect(self.on_config_item_changed)
        del_item_btn = QPushButton("刪除項目"); del_item_btn.clicked.connect(self.delete_config_item)
        layout.addLayout(config_search_layout); layout.addLayout(config_input_layout); layout.addWidget(self.config_table); layout.addWidget(del_item_btn)

    def auto_check_update(self):
        try:
            response = requests.get(UPDATE_URL, timeout=5)
            if response.status_code == 200:
                data = response.json()
                if "equip_data" in data:
                    remote_equip = data["equip_data"]
                    for key in self.equip_data:
                        if key in remote_equip:
                            combined = list(set(self.equip_data[key] + remote_equip[key]))
                            self.equip_data[key] = combined
                    self.refresh_all_combos()
                    self.update_config_table_from_data()
                    self.save_data()

                remote_version = data.get("version", CURRENT_VERSION)
                download_url = data.get("url", "")

                remote_v_list = [int(x) for x in remote_version.split(".")]
                current_v_list = [int(x) for x in CURRENT_VERSION.split(".")]

                if remote_v_list > current_v_list:
                    if not download_url:
                        print(f"DEBUG: 發現新版 {remote_version} 但 JSON 內缺少 'url' 欄位")
                        return

                    reply = QMessageBox.information(
                        self, "發現新版本", 
                        f"目前: v{CURRENT_VERSION}\n最新: v{remote_version}\n\n是否更新？",
                        QMessageBox.Yes | QMessageBox.No
                    )
                    if reply == QMessageBox.Yes:
                        self.execute_auto_update(download_url)
        except Exception as e:
            print(f"檢查更新失敗: {e}")

    def execute_auto_update(self, download_url):
        if not download_url or not download_url.startswith("http"):
            QMessageBox.warning(self, "錯誤", f"下載連結無效或格式錯誤。\n連結內容: {download_url}")
            return
        try:
            current_exe = os.path.abspath(sys.executable)
            new_exe = current_exe + ".new"
            response = requests.get(download_url, stream=True, timeout=30)
            if response.status_code == 200:
                with open(new_exe, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        f.write(chunk)
            else:
                raise Exception(f"伺服器回應錯誤代碼: {response.status_code}")

            bat_path = os.path.join(os.path.dirname(current_exe), "update_helper.bat")
            exe_name = os.path.basename(current_exe)
            
            with open(bat_path, "w", encoding="utf-8") as f:
                f.write(f"""@echo off
taskkill /f /im "{exe_name}" >nul 2>&1
timeout /t 2 /nobreak >nul
del /f /q "{current_exe}"
move /y "{new_exe}" "{current_exe}"
start "" "{current_exe}"
del "%~f0"
""")
            subprocess.Popen([bat_path], shell=True)
            QApplication.quit()
            sys.exit()
        except Exception as e:
            QMessageBox.critical(self, "更新失敗", f"更新過程中發生錯誤：\n{str(e)}")

    def sync_realtime_updates(self):
        if self.start_time:
            elapsed = (time.time() - self.start_time) / 60
            self.timer_btn.setText(f"結束計時 ({elapsed:.2f}m)")

    def toggle_timer(self):
        if self.start_time is None:
            self.start_time = time.time()
            self.timer_display.start(1000)
        else:
            elapsed = (time.time() - self.start_time) / 60
            self.time_input.setValue(elapsed)
            self.start_time = None
            self.timer_btn.setText("開始計時")
            self.timer_display.stop()

    def add_record(self):
        row = self.table.rowCount()
        self.table.insertRow(row)
        self.update_row_from_inputs(row)
        self.save_data()
        self.update_analysis()

    def update_row_from_inputs(self, row):
        diff = self.exp_end_input.value() - self.exp_start_input.value()
        mins = self.time_input.value()
        hr_exp = (diff / mins * 60) if mins > 0 else 0
        data = [
            f"({self.weapon_plus.currentText().replace('+', '')}){self.weapon_input.currentText()}",
            f"({self.neck_plus.currentText().replace('+', '')}){self.neck_input.currentText()}",
            f"({self.card_plus1.currentText().replace('+', '')}){self.card_input1.currentText()}",
            f"({self.card_plus2.currentText().replace('+', '')}){self.card_input2.currentText()}",
            f"({self.card_plus3.currentText().replace('+', '')}){self.card_input3.currentText()}",
            f"({self.card_plus4.currentText().replace('+', '')}){self.card_input4.currentText()}",
            f"({self.mount_plus.currentText().replace('+', '')}){self.mount_input.currentText()}",
            f"({self.soul_plus.currentText().replace('+', '')}){self.soul_input.currentText()}",
            self.pet_input.currentText(), self.crit_input.text(), self.atk_boost_input.text(),
            f"{mins:.2f}", f"{self.exp_start_input.value():.5f}", f"{self.exp_end_input.value():.5f}",
            f"{diff:.5f}", f"{hr_exp:.5f}", self.note_input.text()
        ]
        for c, v in enumerate(data):
            self.table.setItem(row, c, QTableWidgetItem(str(v)))

    def save_data(self):
        all_data = {"records": [], "equip_data": self.equip_data, "next_lvl_exp": self.next_lvl_exp.value()}
        for r in range(self.table.rowCount()):
            all_data["records"].append([self.table.item(r, c).text() for c in range(self.table.columnCount())])
        with open(self.data_file, "w", encoding="utf-8") as f:
            json.dump(all_data, f, ensure_ascii=False, indent=4)

    def load_data(self):
        if os.path.exists(self.data_file):
            try:
                with open(self.data_file, "r", encoding="utf-8") as f:
                    d = json.load(f)
                    self.equip_data.update(d.get("equip_data", {}))
                    self.next_lvl_exp.setValue(d.get("next_lvl_exp", 0))
                    self.refresh_all_combos()
                    self.update_config_table_from_data()
                    for row_data in d.get("records", []):
                        r = self.table.rowCount(); self.table.insertRow(r)
                        for c, v in enumerate(row_data): self.table.setItem(r, c, QTableWidgetItem(v))
            except: pass

    def clear_inputs(self): self.time_input.setValue(0); self.note_input.clear()
    
    def toggle_dark_mode(self): 
        self.is_dark_mode = not self.is_dark_mode
        if self.is_dark_mode: 
            self.setStyleSheet("QMainWindow, QWidget { background-color: #121212; color: #E0E0E0; } QTableWidget { background-color: #1E1E1E; color: white; }")
        else: 
            self.setStyleSheet("")

    def delete_record(self): 
        row = self.table.currentRow()
        if row >= 0: self.table.removeRow(row); self.save_data()

    def copy_last_record(self):
        if self.table.rowCount() > 0: self.load_to_inputs_by_row(self.table.rowCount()-1)

    def load_to_inputs(self, item): self.load_to_inputs_by_row(item.row())

    def load_to_inputs_by_row(self, r):
        """ 將表格選定列的資料載回輸入框 """
        try:
            # 處理帶有 (+數字) 的裝備名稱
            def split_plus(text):
                if text.startswith("(") and ")" in text:
                    parts = text.split(")", 1)
                    return parts[0][1:], parts[1]
                return "0", text

            w_p, w_n = split_plus(self.table.item(r, 0).text())
            self.weapon_plus.setCurrentText(f"+{w_p}")
            self.weapon_input.setCurrentText(w_n)

            n_p, n_n = split_plus(self.table.item(r, 1).text())
            self.neck_plus.setCurrentText(f"+{n_p}")
            self.neck_input.setCurrentText(n_n)

            c1_p, c1_n = split_plus(self.table.item(r, 2).text())
            self.card_plus1.setCurrentText(f"+{c1_p}")
            self.card_input1.setCurrentText(c1_n)

            c2_p, c2_n = split_plus(self.table.item(r, 3).text())
            self.card_plus2.setCurrentText(f"+{c2_p}")
            self.card_input2.setCurrentText(c2_n)

            c3_p, c3_n = split_plus(self.table.item(r, 4).text())
            self.card_plus3.setCurrentText(f"+{c3_p}")
            self.card_input3.setCurrentText(c3_n)

            c4_p, c4_n = split_plus(self.table.item(r, 5).text())
            self.card_plus4.setCurrentText(f"+{c4_p}")
            self.card_input4.setCurrentText(c4_n)

            m_p, m_n = split_plus(self.table.item(r, 6).text())
            self.mount_plus.setCurrentText(f"+{m_p}")
            self.mount_input.setCurrentText(m_n)

            s_p, s_n = split_plus(self.table.item(r, 7).text())
            self.soul_plus.setCurrentText(f"+{s_p}")
            self.soul_input.setCurrentText(s_n)

            self.pet_input.setCurrentText(self.table.item(r, 8).text())
            self.crit_input.setText(self.table.item(r, 9).text())
            self.atk_boost_input.setText(self.table.item(r, 10).text())
            self.time_input.setValue(float(self.table.item(r, 11).text()))
            self.exp_start_input.setValue(float(self.table.item(r, 12).text()))
            self.exp_end_input.setValue(float(self.table.item(r, 13).text()))
            self.note_input.setText(self.table.item(r, 16).text())
        except Exception as e:
            print(f"載入資料失敗: {e}")

    def edit_record(self):
        row = self.table.currentRow()
        if row >= 0:
            self.update_row_from_inputs(row)
            self.save_data()
            self.update_analysis()
            QMessageBox.information(self, "成功", "紀錄已更新")
        else:
            QMessageBox.warning(self, "提示", "請先點選要修改的資料列")

    def filter_table(self):
        search_text = self.search_input.text().lower()
        for r in range(self.table.rowCount()):
            match = False
            for c in range(self.table.columnCount()):
                item = self.table.item(r, c)
                if item and search_text in item.text().lower():
                    match = True
                    break
            self.table.setRowHidden(r, not match)

    def update_analysis(self):
        self.rank_table.setRowCount(0)
        analysis_data = {}
        for r in range(self.table.rowCount()):
            combo = "-".join([self.table.item(r, i).text() for i in range(9)])
            try:
                crit = float(self.table.item(r, 9).text().replace("%", ""))
                hourly_exp = float(self.table.item(r, 15).text())
            except: continue
            if combo not in analysis_data:
                analysis_data[combo] = {"crits": [], "max_hr": 0.0, "count": 0}
            analysis_data[combo]["crits"].append(crit)
            analysis_data[combo]["max_hr"] = max(analysis_data[combo]["max_hr"], hourly_exp)
            analysis_data[combo]["count"] += 1
        sorted_combos = sorted(analysis_data.items(), key=lambda x: x[1]["max_hr"], reverse=True)
        for i, (combo, info) in enumerate(sorted_combos):
            row = self.rank_table.rowCount(); self.rank_table.insertRow(row)
            avg_crit = sum(info["crits"]) / len(info["crits"])
            self.rank_table.setItem(row, 0, QTableWidgetItem(str(i + 1)))
            self.rank_table.setItem(row, 1, QTableWidgetItem(combo))
            self.rank_table.setItem(row, 2, QTableWidgetItem(f"{avg_crit:.2f}%"))
            self.rank_table.setItem(row, 3, QTableWidgetItem(f"{info['max_hr']:.5f}%"))
            self.rank_table.setItem(row, 4, QTableWidgetItem(str(info["count"])))
        self.calculate_countdown()

    def calculate_countdown(self):
        try:
            if self.rank_table.rowCount() > 0:
                best_hr_exp = float(self.rank_table.item(0, 3).text().replace("%", ""))
                remaining = self.next_lvl_exp.value()
                if best_hr_exp > 0:
                    total_hours = remaining / best_hr_exp
                    hours = int(total_hours)
                    minutes = int((total_hours - hours) * 60)
                    self.countdown_label.setText(f"預計升級所需時間: {hours} 小時 {minutes} 分鐘 (以此最高時薪計算)")
                else: self.countdown_label.setText("預計升級所需時間: --")
        except: self.countdown_label.setText("預計升級所需時間: 無效數據")

    def filter_config_table(self):
        text = self.config_search_input.text().lower()
        for r in range(self.config_table.rowCount()):
            match = text in self.config_table.item(r, 0).text().lower() or text in self.config_table.item(r, 1).text().lower()
            self.config_table.setRowHidden(r, not match)

    def on_config_item_changed(self, item): pass

    def add_config_item(self):
        category = self.cate_combo.currentText()
        name = self.item_name_input.text().strip()
        if name:
            if name not in self.equip_data[category]:
                self.equip_data[category].append(name)
                self.update_config_table_from_data(); self.refresh_all_combos(); self.save_data(); self.item_name_input.clear()
            else: QMessageBox.warning(self, "提示", "該項目已存在")

    def delete_config_item(self):
        row = self.config_table.currentRow()
        if row >= 0:
            category = self.config_table.item(row, 0).text()
            name = self.config_table.item(row, 1).text()
            if name in self.equip_data[category]:
                self.equip_data[category].remove(name)
                self.update_config_table_from_data(); self.refresh_all_combos(); self.save_data()
        else: QMessageBox.warning(self, "提示", "請先選擇要刪除的項目")

    def update_config_table_from_data(self):
        self.config_table.setRowCount(0); self.config_table.blockSignals(True)
        for category, items in self.equip_data.items():
            for item in items:
                row = self.config_table.rowCount(); self.config_table.insertRow(row)
                self.config_table.setItem(row, 0, QTableWidgetItem(category))
                self.config_table.setItem(row, 1, QTableWidgetItem(item))
        self.config_table.blockSignals(False)

    def refresh_all_combos(self):
        combos = {"武器": [self.weapon_input], "項鍊": [self.neck_input], "卡片": [self.card_input1, self.card_input2, self.card_input3, self.card_input4],
                  "坐騎": [self.mount_input], "鬥魂": [self.soul_input], "寵物": [self.pet_input]}
        for cat, widgets in combos.items():
            for w in widgets:
                current = w.currentText(); w.clear(); w.addItems(sorted(self.equip_data[cat])); w.setCurrentText(current)

    def import_from_csv(self):
        path, _ = QFileDialog.getOpenFileName(self, "匯入 CSV", "", "CSV Files (*.csv)")
        if path:
            try:
                with open(path, "r", encoding="utf-8-sig") as f:
                    reader = csv.reader(f); next(reader)
                    for row_data in reader:
                        r = self.table.rowCount(); self.table.insertRow(r)
                        for c, v in enumerate(row_data): self.table.setItem(r, c, QTableWidgetItem(v))
                self.save_data(); QMessageBox.information(self, "成功", "匯入完成")
            except Exception as e: QMessageBox.critical(self, "錯誤", f"匯入失敗: {e}")

    def export_to_csv(self):
        path, _ = QFileDialog.getSaveFileName(self, "匯出 CSV", "", "CSV Files (*.csv)")
        if path:
            try:
                with open(path, "w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f)
                    headers = [self.table.horizontalHeaderItem(i).text() for i in range(self.table.columnCount())]
                    writer.writerow(headers)
                    for r in range(self.table.rowCount()):
                        row_data = [self.table.item(r, c).text() for c in range(self.table.columnCount())]
                        writer.writerow(row_data)
                QMessageBox.information(self, "成功", "匯出成功")
            except Exception as e: QMessageBox.critical(self, "錯誤", f"匯出失敗: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = GameTracker()
    window.show()
    sys.exit(app.exec())