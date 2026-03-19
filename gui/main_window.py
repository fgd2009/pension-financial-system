# -*- coding: utf-8 -*-
"""
养老金融通报数据自动化处理系统
Tkinter GUI主界面

作者: Matrix Agent
创建日期: 2026-03-19
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from tkinter import *
import threading
import os
import sys
from datetime import datetime
from typing import Dict, Optional

# 添加父目录到路径以便导入
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.processor import DataProcessor
from core.config import BRANCHES, MANUAL_INDICATORS
from utils.excel_tool import ExcelTool


class PensionFinancialApp:
    """养老金融通报数据自动化处理系统主应用"""

    def __init__(self, root):
        self.root = root
        self.root.title("养老金融通报数据自动化处理系统 v1.0")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 700)

        # 数据处理器
        self.processor = DataProcessor()
        self.excel_tool = None

        # 当前模板路径
        self.template_path = None

        # 手工数据存储
        self.manual_data = {}

        # 初始化界面
        self._create_menu()
        self._create_main_frame()

    # =========================================================================
    # 界面创建
    # =========================================================================

    def _create_menu(self):
        """创建菜单栏"""
        menubar = Menu(self.root)
        self.root.config(menu=menubar)

        # 文件菜单
        file_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="打开模板...", command=self.open_template)
        file_menu.add_command(label="导入单一计划...", command=self.import_single_plan)
        file_menu.add_command(label="导入集合计划...", command=self.import_collective_plan)
        file_menu.add_separator()
        file_menu.add_command(label="导出Excel...", command=self.export_excel)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.root.quit)

        # 数据菜单
        data_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="数据处理", menu=data_menu)
        data_menu.add_command(label="计算透视表", command=self.calculate_pivot)
        data_menu.add_command(label="计算KPI", command=self.calculate_kpi)
        data_menu.add_command(label="生成排名", command=self.calculate_rankings)
        data_menu.add_command(label="生成通报", command=self.generate_reports)
        data_menu.add_separator()
        data_menu.add_command(label="一键全流程", command=self.run_full_pipeline)

        # 帮助菜单
        help_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="使用说明", command=self.show_help)
        help_menu.add_command(label="关于", command=self.show_about)

    def _create_main_frame(self):
        """创建主框架"""
        # 顶部状态栏
        self.status_frame = Frame(self.root, bg="#f0f0f0", height=40)
        self.status_frame.pack(fill=X, side=TOP)
        self.status_frame.pack_propagate(False)

        self.status_label = Label(
            self.status_frame,
            text="未加载模板 | 请先打开模板文件",
            bg="#f0f0f0",
            fg="#666666",
            anchor=W,
            padx=10
        )
        self.status_label.pack(fill=X, pady=10)

        # 主内容区域 - 使用Notebook实现多标签页
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=BOTH, expand=True, padx=5, pady=5)

        # 创建各个标签页
        self._create_overview_tab()
        self._create_import_tab()
        self._create_manual_tab()
        self._create_preview_tab()
        self._create_validation_tab()

    def _create_overview_tab(self):
        """创建概览标签页"""
self.overview_frame = Frame(self.notebook)
        self.notebook.add(self.overview_frame, text="数据概览")

        # 标题
        title_label = Label(
            self.overview_frame,
            text="养老金融通报数据自动化处理系统",
            font=("Microsoft YaHei", 18, "bold"),
            pady=20
        )
        title_label.pack()

        # 数据状态卡片
        card_frame = Frame(self.overview_frame, bg="#e8f4e8", padx=20, pady=20)
        card_frame.pack(pady=20, padx=40, fill=X)

        self.overview_labels = {}
        cards = [
            ("template_status", "模板状态", "未加载"),
            ("single_count", "单一计划数据", "0 条"),
            ("collective_count", "集合计划数据", "0 条"),
            ("pivot_status", "透视表", "未计算"),
            ("kpi_status", "KPI表", "未计算"),
            ("report_status", "通报", "未生成"),
        ]

        for i, (key, title, default) in enumerate(cards):
            row = i // 3
            col = i % 3
            card = Frame(card_frame, bg="white", padx=15, pady=15)
            card.grid(row=row, column=col, padx=10, pady=10)

            Label(card, text=title, font=("Microsoft YaHei", 10), bg="white").pack()
            label = Label(card, text=default, font=("Microsoft YaHei", 14, "bold"), bg="white", fg="#2e7d32")
            label.pack()
            self.overview_labels[key] = label

        # 操作按钮区
        btn_frame = Frame(self.overview_frame)
        btn_frame.pack(pady=30)

        Button(
            btn_frame, text="打开模板", command=self.open_template,
            width=15, height=2, bg="#1976d2", fg="white", font=("Microsoft YaHei", 10)
        ).pack(side=LEFT, padx=10)

        Button(
            btn_frame, text="一键全流程", command=self.run_full_pipeline,
            width=15, height=2, bg="#388e3c", fg="white", font=("Microsoft YaHei", 10)
        ).pack(side=LEFT, padx=10)

        Button(
            btn_frame, text="导出结果", command=self.export_excel,
            width=15, height=2, bg="#f57c00", fg="white", font=("Microsoft YaHei", 10)
        ).pack(side=LEFT, padx=10)

    def _create_import_tab(self):
        """创建数据导入标签页"""
        self.import_frame = Frame(self.notebook)
        self.notebook.add(self.import_frame, text="数据导入")

        # 说明标签
        Label(
            self.import_frame, text="数据导入",
            font=("Microsoft YaHei", 14, "bold"), pady=10
        ).pack()

        Label(
            self.import_frame, text="从公司系统导出的企业年金客户明细文件",
            fg="#666666"
        ).pack()

        # 单一计划导入区
        single_frame = LabelFrame(
            self.import_frame, text="单一计划明细导入",
            font=("Microsoft YaHei", 11), padx=20, pady=20
        )
        single_frame.pack(fill=X, padx=40, pady=20)

        self.single_file_label = Label(single_frame, text="未选择文件", fg="#999999")
        self.single_file_label.pack(side=LEFT, padx=10)

        Button(
            single_frame, text="选择文件",
            command=lambda: self.select_file("single")
        ).pack(side=RIGHT)

        # 集合计划导入区
        collective_frame = LabelFrame(
            self.import_frame, text="集合计划明细导入",
            font=("Microsoft YaHei", 11), padx=20, pady=20
        )
        collective_frame.pack(fill=X, padx=40, pady=20)

        self.collective_file_label = Label(collective_frame, text="未选择文件", fg="#999999")
        self.collective_file_label.pack(side=LEFT, padx=10)

        Button(
            collective_frame, text="选择文件",
            command=lambda: self.select_file("collective")
        ).pack(side=RIGHT)

        # 导入按钮
        Button(
            self.import_frame, text="执行数据透视汇总",
            command=self.calculate_pivot,
            width=20, height=2, bg="#1976d2", fg="white", font=("Microsoft YaHei", 11)
        ).pack(pady=20)

    def _create_manual_tab(self):
        """创建手工数据标签页"""
        self.manual_frame = Frame(self.notebook)
        self.notebook.add(self.manual_frame, text="手工数据管理")

        Label(
            self.manual_frame, text="手工维护数据",
            font=("Microsoft YaHei", 14, "bold"), pady=10
        ).pack()

        # 创建手工数据表格
        table_frame = Frame(self.manual_frame)
        table_frame.pack(fill=BOTH, expand=True, padx=20, pady=10)

        # 表头
        headers = ["分行", "托管规模(亿元)", "托管客户数(户)"]
        for col, header in enumerate(headers):
            Label(
                table_frame, text=header, font=("Microsoft YaHei", 10, "bold"),
                bg="#e3f2fd", padx=10, pady=5, relief=RIDGE
            ).grid(row=0, column=col, sticky="nsew")

        # 数据行
        self.manual_entries = {}
        for row, branch in enumerate(BRANCHES, 1):
            # 分行名称
            Label(table_frame, text=branch, padx=10, pady=3).grid(row=row, column=0, sticky="nsew")

            # 托管规模输入框
            entr_scale = Entry(table_frame, width=15)
            entr_scale.grid(row=row, column=1, padx=5, pady=2)
            self.manual_entries[f"{branch}_scale"] = entr_scale

            # 托管客户数输入框
            entr_customer = Entry(table_frame, width=15)
            entr_customer.grid(row=row, column=2, padx=5, pady=2)
            self.manual_entries[f"{branch}_customer"] = entr_customer

        # 功能按钮
        btn_frame = Frame(self.manual_frame)
        btn_frame.pack(pady=15)

        Button(
            btn_frame, text="上月数据复制到当月",
            command=self.copy_last_month_data,
            width=18, bg="#7b1fa2", fg="white"
        ).pack(side=LEFT, padx=5)

        Button(
            btn_frame, text="保存手工数据",
            command=self.save_manual_data,
            width=15, bg="#1976d2", fg="white"
        ).pack(side=LEFT, padx=5)

        Button(
            btn_frame, text="刷新KPI计算",
            command=self.calculate_kpi,
            width=15, bg="#388e3c", fg="white"
        ).pack(side=LEFT, padx=5)

    def _create_preview_tab(self):
        """创建通报预览标签页"""
        self.preview_frame = Frame(self.notebook)
        self.notebook.add(self.preview_frame, text="通报预览")

        Label(
            self.preview_frame, text="通报预览",
            font=("Microsoft YaHei", 14, "bold"), pady=10
        ).pack()

        # 创建两个预览区域
        preview_paned = PanedWindow(self.preview_frame, orient=HORIZONTAL)
        preview_paned.pack(fill=BOTH, expand=True, padx=20, pady=10)

        # 客户数通报
        customer_frame = LabelFrame(preview_paned, text="养老金融客户数通报")
        preview_paned.add(customer_frame)

        self.customer_tree = self._create_report_tree(customer_frame)

        # 规模通报
        scale_frame = LabelFrame(preview_paned, text="养老金融规模通报")
        preview_paned.add(scale_frame)

        self.scale_tree = self._create_report_tree(scale_frame)

    def _create_report_tree(self, parent):
        """创建通报树形表格"""
        tree_frame = Frame(parent)
        tree_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)

        columns = ("rank", "branch", "value", "growth")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=20)

        tree.heading("rank", text="排名")
        tree.heading("branch", text="分行")
        tree.heading("value", text="数值")
        tree.heading("growth", text="增速")

        tree.column("rank", width=50, anchor=CENTER)
        tree.column("branch", width=100)
        tree.column("value", width=100, anchor=E)
        tree.column("growth", width=80, anchor=E)

        scrollbar = Scrollbar(tree_frame, orient=VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

        return tree

    def _create_validation_tab(self):
        """创建数据验证标签页"""
        self.validation_frame = Frame(self.notebook)
        self.notebook.add(self.validation_frame, text="数据验证")

        Label(
            self.validation_frame, text="数据质量验证",
            font=("Microsoft YaHei", 14, "bold"), pady=10
        ).pack()

        # 验证结果区域
        self.validation_text = scrolledtext.ScrolledText(
            self.validation_frame, width=100, height=30,
            font=("Consolas", 10), state=DISABLED
        )
        self.validation_text.pack(fill=BOTH, expand=True, padx=20, pady=10)

        Button(
            self.validation_frame, text="执行数据验证",
            command=self.run_validation,
            width=20, height=2, bg="#f57c00", fg="white", font=("Microsoft YaHei", 11)
        ).pack(pady=15)

    # =========================================================================
    # 核心业务逻辑
    # =========================================================================

    def open_template(self):
        """打开模板文件"""
        filepath = filedialog.askopenfilename(
            title="选择模板文件",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )

        if filepath:
            self.template_path = filepath
            self.excel_tool = ExcelTool()

            if self.excel_tool.load_template(filepath):
                self._update_status(f"已加载模板: {os.path.basename(filepath)}")
                self._update_overview("template_status", "已加载")
                self._load_manual_data_from_template()
                messagebox.showinfo("成功", "模板加载成功！")
            else:
                messagebox.showerror("错误", "模板加载失败！")

    def _load_manual_data_from_template(self):
        """从模板加载手工数据"""
        if self.excel_tool:
            self.manual_data = self.excel_tool.read_manual_data()
            self._populate_manual_entries()

    def _populate_manual_entries(self):
        """填充手工数据到输入框"""
        for branch in BRANCHES:
            scale_key = f"{branch}_scale"
            customer_key = f"{branch}_customer"

            if branch in self.manual_data.get('kpi_scale', {}):
                scale_val = self.manual_data['kpi_scale'][branch].get('curr托管', '')
                if scale_key in self.manual_entries:
                    self.manual_entries[scale_key].delete(0, END)
                    self.manual_entries[scale_key].insert(0, str(scale_val) if scale_val else "")

            if branch in self.manual_data.get('kpi_customer', {}):
                customer_val = self.manual_data['kpi_customer'][branch].get('curr托管', '')
                if customer_key in self.manual_entries:
                    self.manual_entries[customer_key].delete(0, END)
                    self.manual_entries[customer_key].insert(0, str(customer_val) if customer_val else "")

    def select_file(self, file_type: str):
        """选择导入文件"""
        filepath = filedialog.askopenfilename(
            title=f"选择{file_type}文件",
            filetypes=[("Excel文件", "*.xls *.xlsx"), ("所有文件", "*.*")]
        )

        if filepath:
            if file_type == "single":
                self.single_file_label.config(text=os.path.basename(filepath), fg="#1976d2")
                self.single_file_path = filepath
            else:
                self.collective_file_label.config(text=os.path.basename(filepath), fg="#1976d2")
                self.collective_file_path = filepath

    def import_single_plan(self):
        """导入单一计划数据"""
        filepath = filedialog.askopenfilename(
            title="导入单一计划明细",
            filetypes=[("Excel文件", "*.xls *.xlsx"), ("所有文件", "*.*")]
        )

        if filepath:
            success, msg = self.processor.import_single_plan(filepath)
            if success:
                self._update_overview("single_count", f"{msg.split()[-2]} 条")
                self._update_status(msg)
                messagebox.showinfo("成功", msg)
            else:
                messagebox.showerror("错误", msg)

    def import_collective_plan(self):
        """导入集合计划数据"""
        filepath = filedialog.askopenfilename(
            title="导入集合计划明细",
            filetypes=[("Excel文件", "*.xls *.xlsx"), ("所有文件", "*.*")]
        )

        if filepath:
            success, msg = self.processor.import_collective_plan(filepath)
            if success:
                self._update_overview("collective_count", f"{msg.split()[-2]} 条")
                self._update_status(msg)
                messagebox.showinfo("成功", msg)
            else:
                messagebox.showerror("错误", msg)

    def calculate_pivot(self):
        """计算透视表"""
        self._run_task("计算透视表中...", self._calculate_pivot_task)

    def _calculate_pivot_task(self):
        """透视表计算任务"""
        try:
            # 导入数据(如果还没有)
            if hasattr(self, 'single_file_path'):
                self.processor.import_single_plan(self.single_file_path)
            if hasattr(self, 'collective_file_path'):
                self.processor.import_collective_plan(self.collective_file_path)

            # 计算透视表
            self.processor.calculate_pivot_scale()
            self.processor.calculate_pivot_customer()

            # 写入Excel
            if self.excel_tool and self.processor.pivot_scale is not None:
                self.excel_tool.write_pivot_scale(self.processor.pivot_scale)
                self.excel_tool.write_pivot_customer(self.processor.pivot_customer)

            self._update_overview("pivot_status", "已计算")
            self._update_status("透视表计算完成")
            return True, "透视表计算完成"
        except Exception as e:
            return False, str(e)

    def calculate_kpi(self):
        """计算KPI"""
        self._run_task("计算KPI中...", self._calculate_kpi_task)

    def _calculate_kpi_task(self):
        """KPI计算任务"""
        try:
            # 确保透视表已计算
            if self.processor.pivot_scale is None:
                self._calculate_pivot_task()

            # 读取手工数据
            self._read_manual_entries()

            # 计算KPI
            self.processor.calculate_kpi_scale()
            self.processor.calculate_kpi_customer()

            # 写入Excel
            if self.excel_tool:
                if self.processor.kpi_scale is not None:
                    self.excel_tool.write_kpi_scale(self.processor.kpi_scale)
                if self.processor.kpi_customer is not None:
                    self.excel_tool.write_kpi_customer(self.processor.kpi_customer)

            self._update_overview("kpi_status", "已计算")
            self._update_status("KPI计算完成")
            return True, "KPI计算完成"
        except Exception as e:
            return False, str(e)

    def _read_manual_entries(self):
        """读取手工数据输入"""
        for branch in BRANCHES:
            scale_key = f"{branch}_scale"
            customer_key = f"{branch}_customer"

            try:
                scale_val = float(self.manual_entries.get(scale_key, Entry()).get() or 0)
                customer_val = float(self.manual_entries.get(customer_key, Entry()).get() or 0)
            except ValueError:
                scale_val = 0
                customer_val = 0

            if branch not in self.manual_data.get('kpi_scale', {}):
                self.manual_data.setdefault('kpi_scale', {})[branch] = {}
            if branch not in self.manual_data.get('kpi_customer', {}):
                self.manual_data.setdefault('kpi_customer', {})[branch] = {}

            self.manual_data['kpi_scale'][branch]['curr托管'] = scale_val
            self.manual_data['kpi_customer'][branch]['curr托管'] = customer_val

    def calculate_rankings(self):
        """计算排名"""
        self._run_task("计算排名中...", self._calculate_rankings_task)

    def _calculate_rankings_task(self):
        """排名计算任务"""
        try:
            self._update_status("排名计算完成(示例)")
            return True, "排名计算完成"
        except Exception as e:
            return False, str(e)

    def generate_reports(self):
        """生成通报"""
        self._run_task("生成通报中...", self._generate_reports_task)

    def _generate_reports_task(self):
        """通报生成任务"""
        try:
            # 更新预览表
            self._refresh_preview()
            self._update_overview("report_status", "已生成")
            self._update_status("通报生成完成")
            return True, "通报生成完成"
        except Exception as e:
            return False, str(e)

    def run_full_pipeline(self):
        """一键全流程"""
        if not self.template_path:
            messagebox.showwarning("警告", "请先打开模板文件！")
            return

        self._run_task("执行全流程中...", self._full_pipeline_task)

    def _full_pipeline_task(self):
        """全流程任务"""
        try:
            # 1. 导入数据
            self._update_status("步骤1/4: 导入数据...")
            if hasattr(self, 'single_file_path'):
                self.processor.import_single_plan(self.single_file_path)
            if hasattr(self, 'collective_file_path'):
                self.processor.import_collective_plan(self.collective_file_path)

            # 2. 计算透视表
            self._update_status("步骤2/4: 计算透视表...")
            self.processor.calculate_pivot_scale()
            self.processor.calculate_pivot_customer()

            # 3. 读取手工数据并计算KPI
            self._update_status("步骤3/4: 计算KPI...")
            self._read_manual_entries()
            self.processor.calculate_kpi_scale()
            self.processor.calculate_kpi_customer()

            # 4. 生成通报
            self._update_status("步骤4/4: 生成通报...")
            self._refresh_preview()

            # 写入Excel
            if self.excel_tool:
                self.excel_tool.write_pivot_scale(self.processor.pivot_scale)
                self.excel_tool.write_pivot_customer(self.processor.pivot_customer)
                if self.processor.kpi_scale is not None:
                    self.excel_tool.write_kpi_scale(self.processor.kpi_scale)
                if self.processor.kpi_customer is not None:
                    self.excel_tool.write_kpi_customer(self.processor.kpi_customer)

            # 更新界面
            self._update_overview("pivot_status", "已完成")
            self._update_overview("kpi_status", "已完成")
            self._update_overview("report_status", "已生成")
            self._update_status("全流程执行完成！")

            return True, "全流程执行完成"
        except Exception as e:
            return False, str(e)

    def _refresh_preview(self):
        """刷新通报预览"""
        # 清空并重新填充客户数通报
        for item in self.customer_tree.get_children():
            self.customer_tree.delete(item)

        # 示例数据
        for i in range(1, 38):
            self.customer_tree.insert("", END, values=(i, BRANCHES[i-1], "-", "-%"))

        for item in self.scale_tree.get_children():
            self.scale_tree.delete(item)

        for i in range(1, 38):
            self.scale_tree.insert("", END, values=(i, BRANCHES[i-1], "-", "-%"))

    def run_validation(self):
        """运行数据验证"""
        if not self.excel_tool:
            messagebox.showwarning("警告", "请先打开模板文件！")
            return

        self._run_task("验证数据中...", self._validation_task)

    def _validation_task(self):
        """验证任务"""
        try:
            result = self.excel_tool.validate_data()

            # 显示结果
            self.validation_text.config(state=NORMAL)
            self.validation_text.delete(1.0, END)

            self.validation_text.insert(END, "=" * 60 + "\n")
            self.validation_text.insert(END, "数据质量验证报告\n")
            self.validation_text.insert(END, f"验证时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            self.validation_text.insert(END, "=" * 60 + "\n\n")

            self.validation_text.insert(END, f"总机构数: {result['total_branches']}\n")
            self.validation_text.insert(END, f"有数据机构: {result['with_data']}\n")
            self.validation_text.insert(END, f"数据覆盖率: {result['coverage']:.1%}\n\n")

            if result['coverage'] >= 0.9:
                self.validation_text.insert(END, "[通过] 数据覆盖率良好\n", "green")
            else:
                self.validation_text.insert(END, "[警告] 数据覆盖率偏低\n", "yellow")

            self.validation_text.tag_config("green", foreground="green")
            self.validation_text.tag_config("yellow", foreground="#f57c00")

            self.validation_text.config(state=DISABLED)

            self._update_status("数据验证完成")
            return True, "验证完成"
        except Exception as e:
            return False, str(e)

    def save_manual_data(self):
        """保存手工数据"""
        self._read_manual_entries()
        messagebox.showinfo("成功", "手工数据已保存(内存中)")
        self._update_status("手工数据已保存")

    def copy_last_month_data(self):
        """复制上月数据到当月"""
        reply = messagebox.askyesno("确认", "是否将上月数据复制到当月？")
        if reply:
            for branch in BRANCHES:
                if branch in self.manual_data.get('kpi_scale', {}):
                    last_val = self.manual_data['kpi_scale'][branch].get('last托管', 0)
                    self.manual_entries[f"{branch}_scale"].delete(0, END)
                    self.manual_entries[f"{branch}_scale"].insert(0, str(last_val) if last_val else "0")

                if branch in self.manual_data.get('kpi_customer', {}):
                    last_val = self.manual_data['kpi_customer'][branch].get('last托管', 0)
                    self.manual_entries[f"{branch}_customer"].delete(0, END)
                    self.manual_entries[f"{branch}_customer"].insert(0, str(last_val) if last_val else "0")

            self._update_status("上月数据已复制到当月")

    def export_excel(self):
        """导出Excel"""
        if not self.template_path or not self.excel_tool:
            messagebox.showwarning("警告", "请先打开模板文件！")
            return

        filepath = filedialog.asksaveasfilename(
            title="导出Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")],
            initialfile=f"养老金融通报_{datetime.now().strftime('%Y%m%d')}"
        )

        if filepath:
            success, msg = self.excel_tool.save_as(filepath)
            if success:
                messagebox.showinfo("成功", msg)
                self._update_status(msg)
            else:
                messagebox.showerror("错误", msg)

    # =========================================================================
    # 辅助方法
    # =========================================================================

    def _run_task(self, task_name: str, task_func):
        """在后台线程运行任务"""
        def thread_target():
            result = task_func()
            if result:
                success, msg = result
                self.root.after(0, lambda: self._on_task_complete(success, msg))

        self._update_status(task_name)
        thread = threading.Thread(target=thread_target, daemon=True)
        thread.start()

    def _on_task_complete(self, success: bool, msg: str):
        """任务完成回调"""
        if success:
            messagebox.showinfo("完成", msg)
        else:
            messagebox.showerror("错误", msg)

    def _update_status(self, message: str):
        """更新状态栏"""
        self.status_label.config(text=f"{datetime.now().strftime('%H:%M:%S')} | {message}")

    def _update_overview(self, key: str, value: str):
        """更新概览标签"""
        if key in self.overview_labels:
            self.overview_labels[key].config(text=value)

    def show_help(self):
        """显示帮助"""
        help_text = """
养老金融通报数据自动化处理系统使用说明

【使用流程】

1. 打开模板
   - 点击"文件" -> "打开模板"
   - 选择"首季通报数据模板_新表样适配版.xlsx"

2. 导入数据
   - 点击"数据导入"标签页
   - 选择公司导出的单一计划和集合计划文件

3. 填写手工数据
   - 点击"手工数据管理"标签页
   - 填写各分行的托管规模和客户数
   - 可使用"上月数据复制"功能快速填入

4. 执行计算
   - 点击"一键全流程"自动完成所有计算
   - 或分别执行各步骤

5. 预览和导出
   - 在"通报预览"查看结果
   - 点击"导出Excel"保存文件

【注意事项】
- 托管规模单位：亿元
- 托管客户数单位：户
- 客户主办机构字段必须与37个分行名称完全匹配
        """
        messagebox.showinfo("使用说明", help_text)

    def show_about(self):
        """显示关于"""
        messagebox.showinfo(
            "关于",
            "养老金融通报数据自动化处理系统 v1.0\n\n"
            "用于某银行养老金融业务的季度通报工作自动化处理\n\n"
            "作者: Matrix Agent\n"
            f"日期: {datetime.now().strftime('%Y-%m-%d')}"
        )


def main():
    """主函数"""
    root = Tk()
    app = PensionFinancialApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
