import sys
import os
import json
import csv
import time
from datetime import datetime
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                             QTableWidget, QTableWidgetItem, QHeaderView, QDoubleSpinBox,
                             QMessageBox, QTabWidget, QComboBox, QSizePolicy, QFileDialog)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QIcon, QColor

# 新增繪圖相關庫
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib

# 設定中文字體以防亂碼
matplotlib.rcParams['font.sans-serif'] = ['Microsoft JhengHei']
matplotlib.rcParams['axes.unicode_minus'] = False

# --- 關鍵：處理打包後的檔案路徑 ---
def resource_path(relative_path):
    """ 取得資源絕對路徑，兼容開發環境與 PyInstaller 打包後的路徑 """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

class MpvCanvas(FigureCanvas):
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = fig.add_subplot(111)
        super().__init__(fig)

class GameTracker(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("遊戲練功效率紀錄工具")
        self.resize(1200, 900)
        
        # 資料檔案路徑
        self.data_file = "game_data.json"
        
        # 計時器變數
        self.start_time = None
        self.timer_display = QTimer()
        self.timer_display.timeout.connect(self.update_timer_label)

        # 裝備資料存儲
        self.equip_data = {
            "武器": [], "項鍊": [], "卡片": [], 
            "坐騎": [], "鬥魂": [], "寵物": []
        }

        # 設定視窗圖示
        icon_file = resource_path("Exp.ico")
        if os.path.exists(icon_file):
            self.setWindowIcon(QIcon(icon_file))

        # 主分頁系統
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
        
        # 載入存檔
        self.load_data()
        
        # 預設套用深色模式
        self.is_dark_mode = False
        self.toggle_dark_mode()

    def setup_record_tab(self):
        layout = QVBoxLayout(self.record_tab)
        layout.setSpacing(10)

        plus_levels = [f"+{i}" for i in range(13)]

        def set_resizable(widget):
            widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        # --- 搜尋與工具排 ---
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

        # --- 第一排 ---
        input_layout1 = QHBoxLayout()
        input_layout1.setSpacing(15)
        
        self.weapon_plus = QComboBox(); self.weapon_plus.addItems(plus_levels); self.weapon_plus.setFixedWidth(55)
        self.weapon_input = QComboBox(); set_resizable(self.weapon_input)
        
        self.neck_plus = QComboBox(); self.neck_plus.addItems(plus_levels); self.neck_plus.setFixedWidth(55)
        self.neck_input = QComboBox(); set_resizable(self.neck_input)
        
        self.mount_plus = QComboBox(); self.mount_plus.addItems(plus_levels); self.mount_plus.setFixedWidth(55)
        self.mount_input = QComboBox(); set_resizable(self.mount_input)
        
        self.soul_plus = QComboBox(); self.soul_plus.addItems(plus_levels); self.soul_plus.setFixedWidth(55)
        self.soul_input = QComboBox(); set_resizable(self.soul_input)

        input_layout1.addWidget(QLabel("武器:"))
        input_layout1.addWidget(self.weapon_plus); input_layout1.addWidget(self.weapon_input)
        input_layout1.addWidget(QLabel("項鍊:"))
        input_layout1.addWidget(self.neck_plus); input_layout1.addWidget(self.neck_input)
        input_layout1.addWidget(QLabel("坐騎:"))
        input_layout1.addWidget(self.mount_plus); input_layout1.addWidget(self.mount_input)
        input_layout1.addWidget(QLabel("鬥魂:"))
        input_layout1.addWidget(self.soul_plus); input_layout1.addWidget(self.soul_input)

        # --- 第二排 (卡片/寵物) ---
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
            combo.setMinimumWidth(0)

        input_layout_cards.addWidget(QLabel("卡1:")); input_layout_cards.addWidget(self.card_plus1); input_layout_cards.addWidget(self.card_input1)
        input_layout_cards.addWidget(QLabel("卡2:")); input_layout_cards.addWidget(self.card_plus2); input_layout_cards.addWidget(self.card_input2)
        input_layout_cards.addWidget(QLabel("卡3:")); input_layout_cards.addWidget(self.card_plus3); input_layout_cards.addWidget(self.card_input3)
        input_layout_cards.addWidget(QLabel("卡4:")); input_layout_cards.addWidget(self.card_plus4); input_layout_cards.addWidget(self.card_input4)
        input_layout_cards.addWidget(QLabel("寵物:")); input_layout_cards.addWidget(self.pet_input)

        # --- 第三排 (數值/計時) ---
        input_layout2 = QHBoxLayout()
        self.crit_input = QLineEdit(); set_resizable(self.crit_input); self.crit_input.setPlaceholderText("爆傷 %")
        self.atk_boost_input = QLineEdit(); set_resizable(self.atk_boost_input); self.atk_boost_input.setPlaceholderText("攻擊力增幅")
        
        self.time_input = QDoubleSpinBox(); set_resizable(self.time_input)
        self.time_input.setRange(0.00, 999999); self.time_input.setSuffix(" 分鐘")

        self.exp_start_input = QDoubleSpinBox(); set_resizable(self.exp_start_input)
        self.exp_start_input.setRange(0, 999999999999.99999); self.exp_start_input.setDecimals(5); self.exp_start_input.setSuffix(" 起始")

        self.exp_end_input = QDoubleSpinBox(); set_resizable(self.exp_end_input)
        self.exp_end_input.setRange(0, 999999999999.99999); self.exp_end_input.setDecimals(5); self.exp_end_input.setSuffix(" 結束")
        
        self.note_input = QLineEdit(); set_resizable(self.note_input); self.note_input.setPlaceholderText("備註事項")

        input_layout2.addWidget(QLabel("爆傷:")); input_layout2.addWidget(self.crit_input)
        input_layout2.addWidget(QLabel("攻增:")); input_layout2.addWidget(self.atk_boost_input)
        input_layout2.addWidget(QLabel("時長:")); input_layout2.addWidget(self.time_input)
        input_layout2.addWidget(QLabel("起始:")); input_layout2.addWidget(self.exp_start_input)
        input_layout2.addWidget(QLabel("結束:")); input_layout2.addWidget(self.exp_end_input)
        input_layout2.addWidget(QLabel("備註:")); input_layout2.addWidget(self.note_input)

        # --- 第四排 (按鈕工具排) ---
        btn_layout = QHBoxLayout()
        self.timer_btn = QPushButton("開始計時")
        self.timer_btn.setStyleSheet("background-color: #673AB7; color: white; font-weight: bold; height: 35px;")
        self.timer_btn.clicked.connect(self.toggle_timer)

        self.copy_btn = QPushButton("複製前一筆")
        self.copy_btn.clicked.connect(self.copy_last_record)

        self.clear_btn = QPushButton("一鍵清空")
        self.clear_btn.clicked.connect(self.clear_inputs)

        add_btn = QPushButton("新增紀錄")
        add_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; height: 35px;")
        add_btn.clicked.connect(self.add_record)

        edit_btn = QPushButton("修改選中列")
        edit_btn.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; height: 35px;")
        edit_btn.clicked.connect(self.edit_record)

        del_btn = QPushButton("刪除選中紀錄")
        del_btn.setStyleSheet("background-color: #f44336; color: white; font-weight: bold; height: 35px;")
        del_btn.clicked.connect(self.delete_record)

        btn_layout.addWidget(self.timer_btn)
        btn_layout.addWidget(self.copy_btn)
        btn_layout.addWidget(self.clear_btn)
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(edit_btn)
        btn_layout.addWidget(del_btn)

        # --- 表格 ---
        self.table = QTableWidget()
        self.table.setColumnCount(17)
        self.table.setHorizontalHeaderLabels([
            "武器", "項鍊", "卡片1", "卡片2", "卡片3", "卡片4", "坐騎", "鬥魂", "寵物", 
            "爆傷", "攻增", "時間", "前經驗", "結束經驗", "經驗值", "時薪(Exp/hr)", "備註"
        ])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.itemDoubleClicked.connect(self.load_to_inputs)

        layout.addLayout(tool_layout)
        layout.addLayout(input_layout1)
        layout.addLayout(input_layout_cards)
        layout.addLayout(input_layout2)
        layout.addLayout(btn_layout)
        layout.addWidget(self.table)

    def setup_analysis_tab(self):
        main_layout = QVBoxLayout(self.analysis_tab)
        
        # 上方：升級倒數區
        countdown_group = QHBoxLayout()
        self.next_lvl_exp = QDoubleSpinBox()
        self.next_lvl_exp.setRange(0, 999999999999.99)
        self.next_lvl_exp.setDecimals(5)
        self.next_lvl_exp.setPrefix("距離下級還差: ")
        self.next_lvl_exp.setSuffix(" Exp")
        self.next_lvl_exp.valueChanged.connect(self.calculate_countdown)
        
        self.countdown_label = QLabel("預計升級所需時間: --小時 --分")
        self.countdown_label.setStyleSheet("font-weight: bold; color: #FF9800; font-size: 14px;")
        
        countdown_group.addWidget(self.next_lvl_exp)
        countdown_group.addWidget(self.countdown_label)
        main_layout.addLayout(countdown_group)

        # 中間：排行榜
        info_label = QLabel("### 效率排行榜 (依據時薪 Exp/hr 排序)")
        info_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(info_label)

        self.rank_table = QTableWidget()
        self.rank_table.setColumnCount(5)
        self.rank_table.setHorizontalHeaderLabels(["排名", "關鍵裝備組合", "平均爆傷", "最高時薪", "紀錄次數"])
        self.rank_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        main_layout.addWidget(self.rank_table)

        # 下方：圖表區
        chart_layout = QHBoxLayout()
        self.trend_canvas = MpvCanvas(self, width=5, height=4)
        self.pie_canvas = MpvCanvas(self, width=5, height=4)
        chart_layout.addWidget(self.trend_canvas)
        chart_layout.addWidget(self.pie_canvas)
        main_layout.addLayout(chart_layout)
        
        refresh_rank_btn = QPushButton("刷新分析數據與圖表")
        refresh_rank_btn.setFixedHeight(40)
        refresh_rank_btn.clicked.connect(self.update_analysis)
        main_layout.addWidget(refresh_rank_btn)

    def setup_config_tab(self):
        layout = QVBoxLayout(self.config_tab)
        layout.setSpacing(10)
        
        # --- 搜尋功能排 ---
        config_search_layout = QHBoxLayout()
        self.config_search_input = QLineEdit()
        self.config_search_input.setPlaceholderText("搜尋已存裝備名稱...")
        self.config_search_input.textChanged.connect(self.filter_config_table)
        config_search_layout.addWidget(QLabel("搜尋裝備:"))
        config_search_layout.addWidget(self.config_search_input)
        
        # --- 原本的新增功能排 ---
        config_input_layout = QHBoxLayout()
        self.cate_combo = QComboBox()
        self.cate_combo.addItems(["武器", "項鍊", "卡片", "坐騎", "鬥魂", "寵物"])
        self.cate_combo.currentIndexChanged.connect(self.filter_config_table)
        self.item_name_input = QLineEdit()
        self.item_name_input.setPlaceholderText("輸入裝備名稱")
        self.item_name_input.returnPressed.connect(self.add_config_item)
        
        add_item_btn = QPushButton("新增至選單")
        add_item_btn.clicked.connect(self.add_config_item)
        
        config_input_layout.addWidget(QLabel("分類:"))
        config_input_layout.addWidget(self.cate_combo)
        config_input_layout.addWidget(self.item_name_input)
        config_input_layout.addWidget(add_item_btn)

        self.config_table = QTableWidget()
        self.config_table.setColumnCount(2)
        self.config_table.setHorizontalHeaderLabels(["類別", "名稱"])
        self.config_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # 允許編輯
        self.config_table.itemChanged.connect(self.on_config_item_changed)

        del_item_btn = QPushButton("刪除裝備項目")
        del_item_btn.clicked.connect(self.delete_config_item)

        layout.addLayout(config_search_layout)
        layout.addLayout(config_input_layout)
        layout.addWidget(self.config_table)
        layout.addWidget(del_item_btn)

    # --- 功能邏輯 ---

    def toggle_timer(self):
        if self.start_time is None:
            self.start_time = time.time()
            self.timer_btn.setText("結束計時 (0.00m)")
            self.timer_btn.setStyleSheet("background-color: #FF5722; color: white; font-weight: bold; height: 35px;")
            self.timer_display.start(1000)
        else:
            elapsed = (time.time() - self.start_time) / 60
            self.time_input.setValue(elapsed)
            self.start_time = None
            self.timer_btn.setText("開始計時")
            self.timer_btn.setStyleSheet("background-color: #673AB7; color: white; font-weight: bold; height: 35px;")
            self.timer_display.stop()

    def update_timer_label(self):
        if self.start_time:
            elapsed = (time.time() - self.start_time) / 60
            self.timer_btn.setText(f"結束計時 ({elapsed:.2f}m)")

    def clear_inputs(self):
        self.time_input.setValue(0)
        self.exp_start_input.setValue(0)
        self.exp_end_input.setValue(0)
        self.note_input.clear()

    def copy_last_record(self):
        if self.table.rowCount() > 0:
            last_row = self.table.rowCount() - 1
            self.load_to_inputs_by_row(last_row)
            try:
                prev_end = float(self.table.item(last_row, 13).text())
                self.exp_start_input.setValue(prev_end)
            except: pass

    def add_record(self):
        try:
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.update_row_from_inputs(row)
            self.save_data()
            self.apply_conditional_formatting()
        except Exception as e:
            QMessageBox.critical(self, "錯誤", str(e))

    def edit_record(self):
        row = self.table.currentRow()
        if row >= 0:
            self.update_row_from_inputs(row)
            self.save_data()
            self.apply_conditional_formatting()
        else:
            QMessageBox.warning(self, "提示", "請先點選要修改的資料列")

    def update_row_from_inputs(self, row):
        diff = self.exp_end_input.value() - self.exp_start_input.value()
        minutes = self.time_input.value()
        hourly_exp = (diff / minutes * 60) if minutes > 0 else 0
        
        data = [
            f"({self.weapon_plus.currentText()}){self.weapon_input.currentText()}",
            f"({self.neck_plus.currentText()}){self.neck_input.currentText()}",
            f"({self.card_plus1.currentText()}){self.card_input1.currentText()}",
            f"({self.card_plus2.currentText()}){self.card_input2.currentText()}",
            f"({self.card_plus3.currentText()}){self.card_input3.currentText()}",
            f"({self.card_plus4.currentText()}){self.card_input4.currentText()}",
            f"({self.mount_plus.currentText()}){self.mount_input.currentText()}",
            f"({self.soul_plus.currentText()}){self.soul_input.currentText()}",
            self.pet_input.currentText(),
            self.crit_input.text(), self.atk_boost_input.text(),
            f"{minutes:.2f}", f"{self.exp_start_input.value():.5f}",
            f"{self.exp_end_input.value():.5f}", f"{diff:.5f}", f"{hourly_exp:.5f}",
            self.note_input.text()
        ]
        for col, val in enumerate(data):
            self.table.setItem(row, col, QTableWidgetItem(str(val)))

    def load_to_inputs(self, item):
        self.load_to_inputs_by_row(item.row())

    def load_to_inputs_by_row(self, row):
        def split_plus(text):
            if text.startswith("(+") and ")" in text:
                idx = text.find(")")
                return text[1:idx], text[idx+1:]
            return "+0", text

        w_p, w_v = split_plus(self.table.item(row, 0).text())
        self.weapon_plus.setCurrentText(w_p); self.weapon_input.setCurrentText(w_v)
        n_p, n_v = split_plus(self.table.item(row, 1).text())
        self.neck_plus.setCurrentText(n_p); self.neck_input.setCurrentText(n_v)
        cp1, cv1 = split_plus(self.table.item(row, 2).text())
        self.card_plus1.setCurrentText(cp1); self.card_input1.setCurrentText(cv1)
        cp2, cv2 = split_plus(self.table.item(row, 3).text())
        self.card_plus2.setCurrentText(cp2); self.card_input2.setCurrentText(cv2)
        cp3, cv3 = split_plus(self.table.item(row, 4).text())
        self.card_plus3.setCurrentText(cp3); self.card_input3.setCurrentText(cv3)
        cp4, cv4 = split_plus(self.table.item(row, 5).text())
        self.card_plus4.setCurrentText(cp4); self.card_input4.setCurrentText(cv4)
        m_p, m_v = split_plus(self.table.item(row, 6).text())
        self.mount_plus.setCurrentText(m_p); self.mount_input.setCurrentText(m_v)
        s_p, s_v = split_plus(self.table.item(row, 7).text())
        self.soul_plus.setCurrentText(s_p); self.soul_input.setCurrentText(s_v)

        self.pet_input.setCurrentText(self.table.item(row, 8).text())
        self.crit_input.setText(self.table.item(row, 9).text())
        self.atk_boost_input.setText(self.table.item(row, 10).text())
        self.time_input.setValue(float(self.table.item(row, 11).text()))
        self.exp_start_input.setValue(float(self.table.item(row, 12).text()))
        self.exp_end_input.setValue(float(self.table.item(row, 13).text()))
        self.note_input.setText(self.table.item(row, 16).text())

    def delete_record(self):
        row = self.table.currentRow()
        if row >= 0:
            if QMessageBox.question(self, "確認", "確定刪除？") == QMessageBox.Yes:
                self.table.removeRow(row)
                self.save_data()

    # --- 裝備設定功能 ---

    def add_config_item(self):
        cate = self.cate_combo.currentText()
        name = self.item_name_input.text().strip()
        if name:
            if name not in self.equip_data[cate]:
                self.equip_data[cate].append(name)
                # 暫時阻斷訊號避免觸發 itemChanged
                self.config_table.blockSignals(True)
                row = self.config_table.rowCount()
                self.config_table.insertRow(row)
                item_cate = QTableWidgetItem(cate)
                item_cate.setFlags(item_cate.flags() & ~Qt.ItemIsEditable) # 分類不給改
                self.config_table.setItem(row, 0, item_cate)
                self.config_table.setItem(row, 1, QTableWidgetItem(name))
                self.config_table.blockSignals(False)
                
                self.refresh_all_combos()
                self.filter_config_table() # 新增後套用目前分類過濾
                self.item_name_input.clear()
                self.save_data()

    def on_config_item_changed(self, item):
        # 只有名稱(欄位1)改變時處理
        if item.column() == 1:
            row = item.row()
            cate_item = self.config_table.item(row, 0)
            if cate_item:
                cate = cate_item.text()
                new_name = item.text().strip()
                
                # 重新整理該類別的資料
                new_list = []
                for r in range(self.config_table.rowCount()):
                    if self.config_table.item(r, 0).text() == cate:
                        val = self.config_table.item(r, 1).text().strip()
                        if val: new_list.append(val)
                
                self.equip_data[cate] = list(dict.fromkeys(new_list)) # 去重
                self.refresh_all_combos()
                self.save_data()

    def delete_config_item(self):
        row = self.config_table.currentRow()
        if row >= 0:
            if QMessageBox.question(self, "確認", "確定刪除此裝備項目？") == QMessageBox.Yes:
                cate = self.config_table.item(row, 0).text()
                name = self.config_table.item(row, 1).text()
                
                self.config_table.blockSignals(True)
                self.config_table.removeRow(row)
                self.config_table.blockSignals(False)
                
                # 從資料結構移除
                if name in self.equip_data[cate]:
                    self.equip_data[cate].remove(name)
                
                self.refresh_all_combos()
                self.save_data()

    def filter_config_table(self):
        selected_cate = self.cate_combo.currentText()
        search_text = self.config_search_input.text().lower().strip()
        
        for r in range(self.config_table.rowCount()):
            cate_item = self.config_table.item(r, 0)
            name_item = self.config_table.item(r, 1)
            
            if cate_item and name_item:
                match_cate = (cate_item.text() == selected_cate)
                match_search = (search_text in name_item.text().lower())
                
                if match_cate and match_search:
                    self.config_table.setRowHidden(r, False)
                else:
                    self.config_table.setRowHidden(r, True)

    def refresh_all_combos(self):
        self.update_combo_items(self.weapon_input, "武器")
        self.update_combo_items(self.neck_input, "項鍊")
        self.update_combo_items(self.mount_input, "坐騎")
        self.update_combo_items(self.soul_input, "鬥魂")
        self.update_combo_items(self.pet_input, "寵物")
        for cb in [self.card_input1, self.card_input2, self.card_input3, self.card_input4]:
            self.update_combo_items(cb, "卡片")

    def update_combo_items(self, combo, key):
        current = combo.currentText()
        combo.blockSignals(True)
        combo.clear()
        combo.addItems(self.equip_data[key])
        combo.setCurrentText(current)
        combo.blockSignals(False)

    # --- 資料持久化 ---

    def save_data(self):
        records = []
        for r in range(self.table.rowCount()):
            row_data = [self.table.item(r, c).text() for c in range(self.table.columnCount())]
            records.append(row_data)
        
        full_data = {
            "equip_data": self.equip_data,
            "records": records,
            "next_lvl_exp": self.next_lvl_exp.value()
        }
        with open(self.data_file, "w", encoding="utf-8") as f:
            json.dump(full_data, f, ensure_ascii=False, indent=4)

    def load_data(self):
        if os.path.exists(self.data_file):
            try:
                with open(self.data_file, "r", encoding="utf-8") as f:
                    full_data = json.load(f)
                
                self.equip_data = full_data.get("equip_data", self.equip_data)
                self.next_lvl_exp.setValue(full_data.get("next_lvl_exp", 0))
                self.refresh_all_combos()
                
                # 填入設定表格
                self.config_table.blockSignals(True)
                self.config_table.setRowCount(0)
                for cate, names in self.equip_data.items():
                    for name in names:
                        row = self.config_table.rowCount()
                        self.config_table.insertRow(row)
                        item_cate = QTableWidgetItem(cate)
                        item_cate.setFlags(item_cate.flags() & ~Qt.ItemIsEditable)
                        self.config_table.setItem(row, 0, item_cate)
                        self.config_table.setItem(row, 1, QTableWidgetItem(name))
                self.config_table.blockSignals(False)
                
                # 套用初始過濾
                self.filter_config_table()
                
                # 填入紀錄表格
                records = full_data.get("records", [])
                self.table.setRowCount(0)
                for row_data in records:
                    row = self.table.rowCount()
                    self.table.insertRow(row)
                    for c, val in enumerate(row_data):
                        self.table.setItem(row, c, QTableWidgetItem(str(val)))
                self.apply_conditional_formatting()
                self.update_analysis()
            except: pass

    # --- 分析與工具功能 ---

    def calculate_countdown(self):
        diff_exp = self.next_lvl_exp.value()
        exps = []
        for r in range(self.table.rowCount()):
            try: exps.append(float(self.table.item(r, 15).text()))
            except: pass
        
        if not exps or diff_exp <= 0:
            self.countdown_label.setText("預計升級所需時間: --小時 --分")
            return

        avg_hr = sum(exps) / len(exps)
        if avg_hr <= 0:
            self.countdown_label.setText("預計升級所需時間: 無限")
            return
            
        total_hours = diff_exp / avg_hr
        hrs = int(total_hours)
        mins = int((total_hours - hrs) * 60)
        self.countdown_label.setText(f"預計升級所需時間: {hrs}小時 {mins}分 (依平均時薪)")

    def update_analysis(self):
        analysis = {}
        history_hr = []
        combo_counts = {}

        for r in range(self.table.rowCount()):
            w = self.table.item(r, 0).text()
            n = self.table.item(r, 1).text()
            p = self.table.item(r, 8).text()
            key = f"{w} | {n} | {p}"
            
            try:
                hourly_exp = float(self.table.item(r, 15).text())
                crit_text = self.table.item(r, 9).text().replace("%", "")
                crit = float(crit_text) if crit_text else 0.0
            except: continue
            
            # 排行榜數據
            if key not in analysis:
                analysis[key] = {"total_hr": 0, "count": 0, "crits": [], "max_hr": 0}
            analysis[key]["total_hr"] += hourly_exp
            analysis[key]["count"] += 1
            analysis[key]["crits"].append(crit)
            if hourly_exp > analysis[key]["max_hr"]:
                analysis[key]["max_hr"] = hourly_exp
            
            # 趨勢圖數據
            history_hr.append(hourly_exp)
            
            # 圓餅圖數據
            combo_counts[key] = combo_counts.get(key, 0) + 1

        # 更新排行榜表格
        sorted_analysis = sorted(analysis.items(), key=lambda x: x[1]["max_hr"], reverse=True)
        self.rank_table.setRowCount(0)
        for i, (key, val) in enumerate(sorted_analysis):
            row = self.rank_table.rowCount()
            self.rank_table.insertRow(row)
            avg_crit = sum(val["crits"]) / len(val["crits"])
            self.rank_table.setItem(row, 0, QTableWidgetItem(str(i+1)))
            self.rank_table.setItem(row, 1, QTableWidgetItem(key))
            self.rank_table.setItem(row, 2, QTableWidgetItem(f"{avg_crit:.1f}%"))
            self.rank_table.setItem(row, 3, QTableWidgetItem(f"{val['max_hr']:.2f}"))
            self.rank_table.setItem(row, 4, QTableWidgetItem(str(val["count"])))

        # 繪製趨勢圖
        self.trend_canvas.axes.clear()
        if history_hr:
            self.trend_canvas.axes.plot(range(1, len(history_hr)+1), history_hr, marker='o', color='#2196F3', label='時薪趨勢')
            self.trend_canvas.axes.set_title("時薪 (Exp/hr) 趨勢圖")
            self.trend_canvas.axes.set_xlabel("紀錄筆數")
            self.trend_canvas.axes.set_ylabel("Exp/hr")
            if len(history_hr) > 1:
                # 簡單趨勢線
                z = [i for i in range(len(history_hr))]
                from npy_append_array import npy_append_array # 僅作為暗示，實際不引用外部。用簡單邏輯：
                # 這裡不引入 numpy 以保持純淨，只做點連線
                pass
        self.trend_canvas.draw()

        # 繪製圓餅圖
        self.pie_canvas.axes.clear()
        if combo_counts:
            labels = list(combo_counts.keys())
            sizes = list(combo_counts.values())
            # 標籤太長則縮減
            short_labels = [L[:10]+"..." if len(L)>10 else L for L in labels]
            self.pie_canvas.axes.pie(sizes, labels=short_labels, autopct='%1.1f%%', startangle=140)
            self.pie_canvas.axes.set_title("常用裝備組合分佈")
        self.pie_canvas.draw()
        
        self.calculate_countdown()

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

    def toggle_dark_mode(self):
        if not self.is_dark_mode:
            self.setStyleSheet("""
                QMainWindow, QWidget { background-color: #2b2b2b; color: #ffffff; }
                QTableWidget { background-color: #3c3f41; gridline-color: #555555; color: #ffffff; }
                QHeaderView::section { background-color: #3c3f41; color: #ffffff; }
                QLineEdit, QComboBox, QDoubleSpinBox { background-color: #3c3f41; color: #ffffff; border: 1px solid #555555; }
                QTabWidget::pane { border: 1px solid #555555; }
                QTabBar::tab { background-color: #3c3f41; color: #bbbbbb; padding: 5px; }
                QTabBar::tab:selected { background-color: #4b4b4b; color: #ffffff; }
            """)
            plt.style.use('dark_background')
            self.is_dark_mode = True
        else:
            self.setStyleSheet("")
            plt.style.use('default')
            self.is_dark_mode = False
        self.update_analysis()

    def apply_conditional_formatting(self):
        count = self.table.rowCount()
        if count == 0: return
        
        exps = []
        for r in range(count):
            try:
                exps.append(float(self.table.item(r, 15).text()))
            except: exps.append(0)
            
        avg = sum(exps) / count if count > 0 else 0
        for r in range(count):
            color = QColor(76, 175, 80, 100) if exps[r] >= avg else QColor(244, 67, 54, 100)
            for c in range(self.table.columnCount()):
                item = self.table.item(r, c)
                if item: item.setBackground(color)

    def export_to_csv(self):
        path, _ = QFileDialog.getSaveFileName(self, "匯出資料", "", "CSV Files (*.csv)")
        if path:
            try:
                with open(path, "w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f)
                    headers = [self.table.horizontalHeaderItem(c).text() for c in range(self.table.columnCount())]
                    writer.writerow(headers)
                    for r in range(self.table.rowCount()):
                        row_data = [self.table.item(r, c).text() for c in range(self.table.columnCount())]
                        writer.writerow(row_data)
                QMessageBox.information(self, "成功", "資料已成功匯出至 CSV")
            except Exception as e:
                QMessageBox.critical(self, "錯誤", f"匯出失敗: {str(e)}")

    def import_from_csv(self):
        path, _ = QFileDialog.getOpenFileName(self, "匯入資料", "", "CSV Files (*.csv)")
        if path:
            try:
                with open(path, "r", encoding="utf-8-sig") as f:
                    reader = csv.reader(f)
                    header = next(reader, None) # 跳過標題
                    if header is None: return
                    
                    if QMessageBox.question(self, "確認", "匯入 CSV 將覆蓋目前所有紀錄，是否繼續？", 
                                         QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
                        self.table.setRowCount(0)
                        for row_data in reader:
                            if len(row_data) == self.table.columnCount():
                                row = self.table.rowCount()
                                self.table.insertRow(row)
                                for c, val in enumerate(row_data):
                                    self.table.setItem(row, c, QTableWidgetItem(str(val)))
                        self.save_data()
                        self.apply_conditional_formatting()
                        self.update_analysis()
                        QMessageBox.information(self, "成功", "資料已從 CSV 匯入完成")
            except Exception as e:
                QMessageBox.critical(self, "錯誤", f"匯入失敗: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = GameTracker()
    window.show()
    sys.exit(app.exec())