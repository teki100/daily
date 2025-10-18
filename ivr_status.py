# @Version: 1.3
# @Author: fd
# @Time: 2025/10/18

import sys, os, json
import pandas as pd
from datetime import datetime, timedelta, date, time
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QTextEdit, QFileDialog, QLabel, QHBoxLayout,
    QTreeWidget, QTreeWidgetItem, QDialog, QTableWidget, QMessageBox, QTableWidgetItem, QHeaderView
)
from PyQt6.QtGui import QIcon, QColor, QBrush
from PyQt6.QtCore import Qt

RESOURCE_PATHS = {
    "icon": "note.ico",
    "stylesheet": "light.qss"
}
import os
import sys
import json
from PyQt6.QtWidgets import QMessageBox


class ConfigManager:
    """管理用户配置（班次时间 + 列映射），支持 PyInstaller 打包后使用"""

    def __init__(self, parent=None):
        self.parent = parent  # 可选，用于弹窗提示
        self.config_path = self.get_user_config_path()
        self.config = self.load_config()

    # 路径管理
    def get_user_config_path(self) -> str:
        """返回配置文件路径（跨平台安全）"""
        base_dir = os.path.join(os.path.expanduser("~"), ".daily_report_config")
        os.makedirs(base_dir, exist_ok=True)
        return os.path.join(base_dir, "user_config.json")

    # 默认配置
    def default_config(self) -> dict:
        """生成默认配置"""
        return {
            "shift_config": {
                "A": ["06:00:00", "15:00:00"],
                "B": ["07:00:00", "16:00:00"],
                "C": ["08:00:00", "17:00:00"],
                "D": ["09:00:00", "18:00:00"],
                "E": ["10:00:00", "19:00:00"],
                "F": ["11:00:00", "20:00:00"],
                "G": ["12:00:00", "21:00:00"],
                "H": ["13:00:00", "22:00:00"],
                "I": ["14:00:00", "23:00:00"],
                "J": ["15:00:00", "00:00:00"],
                "K": ["16:00:00", "01:00:00"],
                "L": ["17:00:00", "02:00:00"]
            },
            "column_config": {
                "表名": "Sheet1",
                "日期": "日期",
                "姓名": "姓名",
                "班次": "班次",
                "状态": "状态",
                "开始时间": "开始时间",
                "结束时间": "结束时间",
                "持续时长min": "持续时长min"
            }
        }

    # 加载与保存
    def load_config(self) -> dict:
        """加载配置文件（不存在则创建默认配置）"""
        if not os.path.exists(self.config_path):
            config = self.default_config()
            self.save_config(config)
            return config
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                config = json.load(f)
            return config
        except Exception as e:
            QMessageBox.warning(self.parent, "加载失败", str(e))
            return self.default_config()

    def save_config(self, config: dict | None = None):
        """保存配置文件"""
        if config is None:
            config = self.config
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
        except Exception as e:
            QMessageBox.warning(self.parent, "保存失败", str(e))

    # 获取配置项
    def get_shift_config(self) -> dict:
        return self.config.get("shift_config", {})

    def get_column_config(self) -> dict:
        return self.config.get("column_config", {})

    # 更新配置项
    def update_shift(self, shift_dict: dict):
        self.config["shift_config"] = shift_dict
        self.save_config()

    def update_columns(self, column_dict: dict):
        self.config["column_config"] = column_dict
        self.save_config()


# 工具函数
def resource_path(relative_path):
    """获取资源文件路径（兼容打包/开发模式）"""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


def load_stylesheet():
    # 加载样式表（只加载一次）
    try:
        with open(resource_path(RESOURCE_PATHS["stylesheet"]), "r", encoding="utf-8") as f:
            return f.read()
    except Exception:
        return ""


# 预加载样式表
STYLESHEET = load_stylesheet()


def force_datetime(x):
    """将各种时间类型统一转成 datetime"""
    if pd.isna(x):
        return pd.NaT
    if isinstance(x, datetime):
        return x
    if isinstance(x, time):
        return datetime.combine(date.today(), x)
    try:
        return pd.to_datetime(x)
    except Exception:
        return pd.NaT


class ChangeDialog(QDialog):
    def __init__(self, config_manager, parent=None):
        super().__init__(parent)
        self.setWindowTitle("修改配置")
        self.setFixedSize(250, 200)
        self.config_manager = config_manager

        # 创建导入和导出的按钮
        self.btn_shift = QPushButton("修改班次时间")
        self.btn_column = QPushButton("修改表格列名")
        self.btn_cancel = QPushButton("⛌  取消")
        self.btn_cancel.clicked.connect(self.close)

        self.btn_shift.clicked.connect(self.shift_change)
        self.btn_column.clicked.connect(self.column_change)

        # 布局
        button_layout = QVBoxLayout()
        button_layout.addWidget(self.btn_shift)
        button_layout.addWidget(self.btn_column)
        button_layout.addWidget(self.btn_cancel)

        self.setLayout(button_layout)

    def shift_change(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("编辑班次时间")
        dialog.resize(400, 490)

        vbox = QVBoxLayout(dialog)
        table = QTableWidget(dialog)
        table.verticalHeader().setVisible(False)

        shift_dict = self.config_manager.get_shift_config()
        table.setRowCount(len(shift_dict))
        table.setColumnCount(3)
        table.setHorizontalHeaderLabels(["班次", "开始时间", "结束时间"])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

        for i, (k, v) in enumerate(shift_dict.items()):
            item_field = QTableWidgetItem(k)
            item_field.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
            table.setItem(i, 0, item_field)
            item_field.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

            item_value1 = QTableWidgetItem(v[0])
            item_value1.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            table.setItem(i, 1, item_value1)
            item_value2 = QTableWidgetItem(v[1])
            item_value2.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            table.setItem(i, 2, item_value2)

        vbox.addWidget(table)
        btn_layout = QHBoxLayout()
        btn_save = QPushButton("保存修改")
        btn_cancel = QPushButton("⛌  取消")
        btn_cancel.clicked.connect(dialog.close)
        btn_layout.addWidget(btn_save)
        btn_layout.addWidget(btn_cancel)
        vbox.addLayout(btn_layout)

        def save_shift():
            current_item = table.currentItem()
            if current_item is not None:
                table.closePersistentEditor(current_item)
            new_dict = {}
            for i in range(table.rowCount()):
                shift = table.item(i, 0).text().strip()
                start = table.item(i, 1).text().strip()
                end = table.item(i, 2).text().strip()
                if shift and start and end:
                    new_dict[shift] = [start, end]
            self.config_manager.update_shift(new_dict)
            QMessageBox.warning(self, "成功", "班次时间已更新！")
            dialog.close()

        btn_save.clicked.connect(save_shift)
        dialog.exec()

    # ================= 列名修改 =================
    def column_change(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("编辑列名")
        dialog.resize(400, 370)

        vbox = QVBoxLayout(dialog)
        table = QTableWidget(dialog)
        table.verticalHeader().setVisible(False)

        col_dict = self.config_manager.get_column_config()
        table.setRowCount(len(col_dict))
        table.setColumnCount(2)
        table.setHorizontalHeaderLabels(["字段", "列名"])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

        for i, (k, v) in enumerate(col_dict.items()):
            item_field = QTableWidgetItem(k)
            item_field.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
            item_field.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            table.setItem(i, 0, item_field)

            item_value = QTableWidgetItem(v)
            item_value.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            table.setItem(i, 1, item_value)

        vbox.addWidget(table)
        btn_layout = QHBoxLayout()
        btn_save = QPushButton("保存修改")
        btn_cancel = QPushButton("⛌  取消")
        btn_cancel.clicked.connect(dialog.reject)
        btn_layout.addWidget(btn_save)
        btn_layout.addWidget(btn_cancel)
        vbox.addLayout(btn_layout)

        def save_column():
            current_item = table.currentItem()
            if current_item is not None:
                table.closePersistentEditor(current_item)
            new_dict = {}
            for i in range(table.rowCount()):
                field = table.item(i, 0).text().strip()
                col = table.item(i, 1).text().strip()
                if field and col:
                    new_dict[field] = col
            print("要保存的列名配置:", new_dict)
            self.config_manager.update_columns(new_dict)
            QMessageBox.warning(self, "成功", "列名配置已更新！")
            dialog.close()

        btn_save.clicked.connect(save_column)
        dialog.exec()


class DailyReportApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("日报生成器")
        self.resize(700, 600)
        self.setMinimumSize(500, 400)
        self.setWindowIcon(QIcon(resource_path(RESOURCE_PATHS["icon"])))

        self.config_manager = ConfigManager(parent=self)
        self.shift_dict = self.config_manager.get_shift_config()
        self.column_map = self.config_manager.get_column_config()

        layout = QVBoxLayout(self)
        top_area = QHBoxLayout()
        layout.addLayout(top_area)

        self.btn_select = QPushButton("选择文件")
        self.btn_select.setFixedWidth(100)
        self.btn_select.clicked.connect(self.select_file)
        top_area.addWidget(self.btn_select)

        self.label = QLabel("请选择 Excel 文件：")
        top_area.addWidget(self.label)

        self.btn_modify = QPushButton("")
        self.btn_modify.setIcon(QIcon(resource_path("setting.svg")))
        self.btn_modify.setFixedWidth(40)
        self.btn_modify.clicked.connect(self.change_dialog)
        top_area.addWidget(self.btn_modify)

        # QTreeWidget 显示报告
        content_area = QHBoxLayout()
        layout.addLayout(content_area)
        self.daily_tree = QTreeWidget()
        self.summary_tree = QTreeWidget()
        for tree in (self.daily_tree, self.summary_tree):
            tree.setColumnCount(1)
            tree.header().hide()
            content_area.addWidget(tree)

        self.file_path = None

        self.daily_tree.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.daily_tree.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.summary_tree.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.summary_tree.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        def enable_click_expand(tree: QTreeWidget):
            def on_item_clicked(item: QTreeWidgetItem):
                if item.childCount() > 0:
                    # 如果节点已展开，则折叠；否则展开
                    item.setExpanded(not item.isExpanded())

            tree.itemClicked.connect(on_item_clicked)

        # 调用
        enable_click_expand(self.summary_tree)
        enable_click_expand(self.daily_tree)

    def change_dialog(self):
        dlg = ChangeDialog(self.config_manager, self)
        dlg.exec()
        # 更新配置
        self.shift_dict = self.config_manager.get_shift_config()
        self.column_map = self.config_manager.get_column_config()

    def select_file(self):
        """选择 Excel 文件并自动生成日报"""
        file_name, _ = QFileDialog.getOpenFileName(
            self, "选择 Excel 文件", "", "Excel Files (*.xlsx *.xls)"
        )
        if not file_name:
            return

        self.file_path = file_name
        self.label.setText(f"已选择文件：{file_name}")
        self.daily_tree.clear()
        self.summary_tree.clear()
        try:
            self._report_content(file_name)
        except Exception as e:
            self.label.setText(f"生成失败：{str(e)}")
        QApplication.processEvents()  # 刷新界面

    def load_excel_data(self, file_name):
        try:
            xls = pd.ExcelFile(file_name)
            first_sheet = xls.sheet_names[0]  # 永远读取第一个 sheet
            df_data = pd.read_excel(xls, sheet_name=first_sheet)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取 Excel 文件失败:\n{e}")
            return pd.DataFrame()

        column_map = self.config_manager.get_column_config()

        rename_dict = {}
        for display_name, excel_name in column_map.items():
            if excel_name in df_data.columns:
                rename_dict[excel_name] = display_name
            else:
                df_data[display_name] = None

        df_data = df_data.rename(columns=rename_dict)

        return df_data

    def _report_content(self, data_file):
        report_dict = {}

        df_data = self.load_excel_data(data_file)
        df_shift = pd.DataFrame([
            {"班次": k, "开始时间": v[0], "结束时间": v[1]}
            for k, v in self.shift_dict.items()
        ])

        # 转换时间列为 datetime
        for col in ["开始时间", "结束时间"]:
            df_data[col] = df_data[col].apply(force_datetime)
            df_shift[col] = df_shift[col].apply(force_datetime)

        df_data["持续时长min"] = pd.to_numeric(df_data["持续时长min"], errors="coerce").round(2)
        df_data["日期"] = pd.to_datetime(df_data["日期"], errors="coerce")

        # 遍历记录
        for idx, row in df_data.iterrows():
            date_str = row["日期"]
            emp = row["姓名"]
            shift = row["班次"]
            status = row["状态"]
            start_t = row["开始时间"]
            end_t = row["结束时间"]
            dur = row["持续时长min"]

            if pd.isna(start_t) or pd.isna(end_t) or pd.isna(shift) or pd.isna(date_str):
                continue

            shift_row = df_shift[df_shift["班次"] == shift]
            if shift_row.empty:
                continue
            shift_start = shift_row.iloc[0]["开始时间"]
            shift_end = shift_row.iloc[0]["结束时间"]
            if shift_end < shift_start:
                shift_end += timedelta(days=1)

            # 将班次时间移到同一天的时间线上
            shift_start = datetime.combine(start_t.date(), shift_start.time())
            shift_end = datetime.combine(start_t.date(), shift_end.time())
            if shift_end < shift_start:
                shift_end += timedelta(days=1)
            if status == "小休" and start_t >= shift_end:
                continue

            if date_str not in report_dict:
                report_dict[date_str] = {
                    "meal": {},
                    "break_sum": {},
                    "break_once": {},
                    "break_shift": {}
                }

            # 用餐统计
            if status == "就餐":
                report_dict[date_str]["meal"][emp] = report_dict[date_str]["meal"].get(emp, 0) + dur

            # 小休总时长
            if status == "小休":
                report_dict[date_str]["break_sum"][emp] = report_dict[date_str]["break_sum"].get(emp, 0) + dur

            # 单次小休超过8分钟
            if status == "小休" and dur > 8:
                report_dict[date_str]["break_once"][emp] = report_dict[date_str]["break_once"].get(emp, 0) + 1

            # 首尾半小时小休
            if status == "小休":
                if (shift_start < start_t < shift_start + timedelta(minutes=30)) or \
                        (shift_end - timedelta(minutes=30) < start_t < shift_end):
                    report_dict[date_str]["break_shift"][emp] = report_dict[date_str]["break_shift"].get(emp, 0) + 1

        # 清空树
        self.daily_tree.clear()
        # 生成树节点
        for date_str in sorted(report_dict.keys()):
            day_item = QTreeWidgetItem(self.daily_tree)
            day_item.setText(0, f"{date_str}")

            # 用餐
            meal_dict = report_dict[date_str]["meal"]
            meal_over = {k: v for k, v in meal_dict.items() if v > 46}
            meal_item = QTreeWidgetItem(day_item)
            if not meal_over:
                meal_item.setText(0, "用餐时长：全员遵时")
            else:
                meal_item.setText(0, f"用餐时长：{len(meal_over)}人超时")
                for k, v in meal_over.items():
                    sub = QTreeWidgetItem(meal_item)
                    sub.setText(0, f"{k}：{v:.1f} 分钟")
                    sub.setForeground(0, QBrush(QColor("#ED856E")))

            # 小休总时长
            break_sum_dict = report_dict[date_str]["break_sum"]
            break_over = {k: v for k, v in break_sum_dict.items() if v > 61}
            break_item = QTreeWidgetItem(day_item)
            if not break_over:
                break_item.setText(0, "小休总时长：全员遵时")
            else:
                break_item.setText(0, f"小休总时长：{len(break_over)}人超时")
                for k, v in break_over.items():
                    sub = QTreeWidgetItem(break_item)
                    sub.setText(0, f"{k}：{v:.1f} 分钟")
                    sub.setForeground(0, QBrush(QColor("#ED856E")))

            # 单次小休
            break_once_dict = report_dict[date_str]["break_once"]
            once_item = QTreeWidgetItem(day_item)
            if not break_once_dict:
                once_item.setText(0, "单次小休：全员遵时")
            else:
                sorted_once = sorted(break_once_dict.items(), key=lambda x: x[1], reverse=True)
                total_people = len(sorted_once)
                total_times = sum(v for _, v in sorted_once)
                once_item.setText(0, f"单次小休：{total_people}人超时，共{total_times}次")
                top_cnt = sorted_once[0][1]
                top_names = [k for k, v in sorted_once if v == top_cnt]
                top_sub = QTreeWidgetItem(once_item)
                top_sub.setText(0, f"小休超时Top：{', '.join(top_names)} {top_cnt}次")
                top_sub.setForeground(0, QBrush(QColor("#ED856E")))
                for k, v in sorted_once:
                    sub = QTreeWidgetItem(once_item)
                    sub.setText(0, f"{k}：{v} 次")
                    sub.setForeground(0, QBrush(QColor("#ED856E")))

            # 首尾半小时小休
            break_shift_dict = report_dict[date_str]["break_shift"]
            shift_item = QTreeWidgetItem(day_item)
            if not break_shift_dict:
                shift_item.setText(0, "首尾半小时小休：全员遵时")
            else:
                sorted_shift = sorted(break_shift_dict.items(), key=lambda x: x[1], reverse=True)
                total_people = len(sorted_shift)
                total_times = sum(v for _, v in sorted_shift)
                shift_item.setText(0, f"首尾半小时小休：{total_people}人，共{total_times}次")
                for k, v in sorted_shift:
                    sub = QTreeWidgetItem(shift_item)
                    sub.setText(0, f"{k}：{v} 次")
                    sub.setForeground(0, QBrush(QColor("#ED856E")))

            day_item.setExpanded(False)  # 默认折叠

        self.daily_tree.expandToDepth(0)

        # 生成 summary 树
        self.summary_tree.clear()

        if report_dict:
            dates = sorted(report_dict.keys())
            start_date = dates[0]
            end_date = dates[-1]

            root_item = QTreeWidgetItem(self.summary_tree)
            root_item.setText(0, f"日期范围：{start_date} - {end_date}")

            # 用餐统计
            meal_total = {}
            for day in report_dict.values():
                for k, v in day.get("meal", {}).items():
                    if v > 46:
                        meal_total[k] = meal_total.get(k, 0) + 1

            meal_item = QTreeWidgetItem(root_item)
            if not meal_total:
                meal_item.setText(0, "用餐时长：全员遵时")
            else:
                meal_item.setText(0, f"用餐时长：{len(meal_total)}人超时，共{sum(meal_total.values())}次")
                for k, v in sorted(meal_total.items(), key=lambda x: x[1], reverse=True):
                    sub = QTreeWidgetItem(meal_item)
                    sub.setText(0, f"{k}：{v} 次")
                    sub.setForeground(0, QBrush(QColor("#ED856E")))

            # 小休总时长
            break_total = {}
            for day in report_dict.values():
                for k, v in day.get("break_sum", {}).items():
                    if v > 61:
                        break_total[k] = break_total.get(k, 0) + 1

            break_item = QTreeWidgetItem(root_item)
            if not break_total:
                break_item.setText(0, "小休总时长：全员遵时")
            else:
                break_item.setText(0, f"小休总时长：{len(break_total)}人超时，共{sum(break_total.values())}次")
                for k, v in sorted(break_total.items(), key=lambda x: x[1], reverse=True):
                    sub = QTreeWidgetItem(break_item)
                    sub.setText(0, f"{k}：{v} 次")
                    sub.setForeground(0, QBrush(QColor("#ED856E")))

            # 单次小休
            once_total = {}
            for day in report_dict.values():
                for k, v in day.get("break_once", {}).items():
                    once_total[k] = once_total.get(k, 0) + v

            once_item = QTreeWidgetItem(root_item)
            if not once_total:
                once_item.setText(0, "单次小休：全员遵时")
            else:
                once_item.setText(0, f"单次小休：{len(once_total)}人超时，共{sum(once_total.values())}次")
                for k, v in sorted(once_total.items(), key=lambda x: x[1], reverse=True):
                    sub = QTreeWidgetItem(once_item)
                    sub.setText(0, f"{k}：{v} 次")
                    sub.setForeground(0, QBrush(QColor("#ED856E")))

            # 首尾半小时小休
            shift_total = {}
            for day in report_dict.values():
                for k, v in day.get("break_shift", {}).items():
                    shift_total[k] = shift_total.get(k, 0) + v

            shift_item = QTreeWidgetItem(root_item)
            if not shift_total:
                shift_item.setText(0, "首尾半小时小休：全员遵时")
            else:
                shift_item.setText(0, f"首尾半小时小休：{len(shift_total)}人，共{sum(shift_total.values())}次")
                for k, v in sorted(shift_total.items(), key=lambda x: x[1], reverse=True):
                    sub = QTreeWidgetItem(shift_item)
                    sub.setText(0, f"{k}：{v} 次")
                    sub.setForeground(0, QBrush(QColor("#ED856E")))

            root_item.setExpanded(False)
            self.summary_tree.expandToDepth(0)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(STYLESHEET)
    window = DailyReportApp()
    window.show()
    sys.exit(app.exec())
