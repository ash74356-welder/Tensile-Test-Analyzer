import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib import rcParams, font_manager
import io
import os
import json
from tkinter import simpledialog

# 设置matplotlib全局字体 - 使用系统字体
def set_matplotlib_font():
    """设置matplotlib中文字体为宋体，英文字体为Times New Roman"""
    try:
        # 查找系统字体
        system_fonts = font_manager.findSystemFonts()
        
        # 寻找宋体
        simsun_found = False
        times_found = False
        
        for font_path in system_fonts:
            font_name = font_manager.FontProperties(fname=font_path).get_name()
            font_family = font_manager.FontProperties(fname=font_path).get_family()
            
            # 寻找宋体或类似中文字体
            if 'SimSun' in font_name or '宋体' in font_name or 'Song' in font_name.lower():
                rcParams['font.sans-serif'] = [font_name, 'DejaVu Sans']
                simsun_found = True
                print(f"找到中文字体: {font_name}")
            
            # 寻找Times New Roman
            if 'Times New Roman' in font_name or 'Times' in font_name:
                rcParams['mathtext.default'] = 'regular'
                rcParams['mathtext.fontset'] = 'stix'
                times_found = True
                print(f"找到英文字体: {font_name}")
        
        # 如果没有找到特定字体，使用通用设置
        if not simsun_found:
            rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'DejaVu Sans']
        
        rcParams['axes.unicode_minus'] = False
        
    except Exception as e:
        print(f"字体设置出错: {e}")
        # 使用默认设置
        rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'DejaVu Sans']
        rcParams['axes.unicode_minus'] = False

# 初始化matplotlib字体
set_matplotlib_font()

class TensileTestAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("铍镍铜拉伸测试数据分析")
        self.root.geometry("1400x900")
        
        # 设置全局字体大小
        self.font_large = ("宋体", 12)
        self.font_medium = ("宋体", 11)
        self.font_small = ("宋体", 10)
        self.font_title = ("宋体", 16, "bold")
        
        # 测试参数
        self.cross_sectional_areas = {}  # 存储每个sheet的横截面积
        self.gauge_length = 10.0  # 引伸计标距 (mm)
        
        # 数据存储
        self.data = None
        self.excel_data = {}  # 存储从Excel读取的所有sheet数据
        self.current_sheet_name = None  # 当前选中的sheet名称
        self.current_excel_path = None  # 当前加载的Excel文件路径
        
        # 图例文本存储
        self.legend_texts = {}
        
        # 配置文件路径
        self.config_file = "tensile_test_config.json"
        
        # 加载配置
        self.load_config()
        
        # 配置样式
        self.setup_styles()
        self.setup_ui()
        
        # 绑定窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        
    def setup_styles(self):
        """设置ttk样式"""
        style = ttk.Style()
        
        # 设置大字体
        style.configure("Title.TLabel", font=self.font_title)
        style.configure("Large.TLabel", font=self.font_large)
        style.configure("Medium.TLabel", font=self.font_medium)
        style.configure("Small.TLabel", font=self.font_small)
        
        style.configure("Large.TButton", font=self.font_large)
        style.configure("Medium.TButton", font=self.font_medium)
        
        style.configure("Large.TEntry", font=self.font_large)
        style.configure("Large.TCombobox", font=self.font_large)
        
        # 设置LabelFrame的字体
        style.configure("TLabelframe", font=self.font_large)
        style.configure("TLabelframe.Label", font=self.font_large)
        
    def setup_ui(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.columnconfigure(2, weight=2)
        main_frame.rowconfigure(2, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="铍镍铜拉伸测试数据分析系统", 
                               style="Title.TLabel")
        title_label.grid(row=0, column=0, columnspan=4, pady=(0, 25))
        
        # 参数输入区域 - 改为动态创建
        self.param_frame = ttk.LabelFrame(main_frame, text="测试参数输入（按Sheet设置）", padding="15")
        self.param_frame.grid(row=1, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(0, 20))
        
        # 参数输入区域初始提示
        self.param_label = ttk.Label(self.param_frame, text="加载Excel文件后，将在此显示各Sheet的横截面积输入框", 
                                    style="Medium.TLabel")
        self.param_label.pack()
        
        # Excel sheet选择区域
        sheet_frame = ttk.LabelFrame(main_frame, text="Excel数据选择", padding="15")
        sheet_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 15), pady=(0, 15))
        
        ttk.Label(sheet_frame, text="选择Sheet:", style="Large.TLabel").grid(row=0, column=0, padx=(0, 15))
        self.sheet_combobox = ttk.Combobox(sheet_frame, width=30, style="Large.TCombobox", state="readonly")
        self.sheet_combobox.grid(row=0, column=1, padx=(0, 20))
        self.sheet_combobox.bind("<<ComboboxSelected>>", self.on_sheet_select)
        
        ttk.Button(sheet_frame, text="加载Excel数据", command=self.load_excel_data, 
                  style="Large.TButton").grid(row=0, column=2, padx=(10, 0))
        
        # 数据预览区域
        preview_frame = ttk.LabelFrame(main_frame, text="数据预览", padding="15")
        preview_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 15))
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        
        # 数据预览文本框 - 使用更大的字体
        self.preview_text = tk.Text(preview_frame, height=20, width=70, 
                                   font=("宋体", 11), wrap=tk.NONE)
        self.preview_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 预览滚动条
        preview_scrollbar = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview_text.yview)
        preview_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.preview_text.configure(yscrollcommand=preview_scrollbar.set)
        
        # 水平滚动条
        preview_h_scrollbar = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.preview_text.xview)
        preview_h_scrollbar.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E))
        self.preview_text.configure(xscrollcommand=preview_h_scrollbar.set)
        
        # 预览信息标签
        self.preview_info_label = ttk.Label(preview_frame, text="未加载数据", style="Medium.TLabel")
        self.preview_info_label.grid(row=2, column=0, columnspan=2, pady=(10, 0))
        
        # 结果显示区域
        result_frame = ttk.LabelFrame(main_frame, text="计算结果", padding="15")
        result_frame.grid(row=3, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 15))
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)
        
        # 结果文本框 - 使用更大的字体
        self.results_text = tk.Text(result_frame, height=10, width=50, 
                                   font=("宋体", 11), state='disabled')
        self.results_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # 多sheet结果显示
        self.multi_results_text = tk.Text(result_frame, height=10, width=50, 
                                         font=("宋体", 11), state='disabled')
        self.multi_results_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 图形显示区域
        plot_frame = ttk.LabelFrame(main_frame, text="载荷-位移曲线", padding="15")
        plot_frame.grid(row=2, column=2, rowspan=2, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        plot_frame.columnconfigure(0, weight=1)
        plot_frame.rowconfigure(0, weight=1)
        
        # 创建图形，设置字体
        self.fig, self.ax = plt.subplots(figsize=(10, 7))
        
        # 设置图形字体
        self.set_plot_font()
        
        self.canvas = FigureCanvasTkAgg(self.fig, master=plot_frame)
        self.canvas.get_tk_widget().grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 图例文本编辑按钮
        legend_frame = ttk.Frame(plot_frame)
        legend_frame.grid(row=1, column=0, pady=(10, 0), sticky=(tk.W, tk.E))
        
        ttk.Button(legend_frame, text="编辑图例文本", command=self.edit_legend_texts,
                  style="Medium.TButton").pack(side=tk.LEFT, padx=5)
        ttk.Button(legend_frame, text="重置图例", command=self.reset_legend_texts,
                  style="Medium.TButton").pack(side=tk.LEFT, padx=5)
        
        # 底部按钮
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.grid(row=4, column=0, columnspan=4, pady=(25, 0))
        
        ttk.Button(bottom_frame, text="处理当前Sheet数据", command=self.process_current_sheet, 
                  style="Large.TButton", width=20).pack(side=tk.LEFT, padx=10)
        ttk.Button(bottom_frame, text="批量处理所有Sheet", command=self.process_all_sheets, 
                  style="Large.TButton", width=20).pack(side=tk.LEFT, padx=10)
        ttk.Button(bottom_frame, text="保存图表", command=self.save_plot, 
                  style="Large.TButton", width=15).pack(side=tk.LEFT, padx=10)
        ttk.Button(bottom_frame, text="导出所有结果", command=self.export_all_results, 
                  style="Large.TButton", width=15).pack(side=tk.LEFT, padx=10)
        ttk.Button(bottom_frame, text="退出程序", command=self.root.quit, 
                  style="Large.TButton", width=10).pack(side=tk.LEFT, padx=10)
    
    def set_plot_font(self):
        """设置图形字体"""
        try:
            # 设置字体
            plt.rcParams['font.sans-serif'] = ['SimSun', 'Times New Roman']
            plt.rcParams['axes.unicode_minus'] = False
            
            # 如果没有SimSun字体，尝试其他字体
            available_fonts = [f.name for f in font_manager.fontManager.ttflist]
            if 'SimSun' not in available_fonts:
                # 尝试其他中文字体
                for font in ['Microsoft YaHei', 'SimHei', 'DejaVu Sans']:
                    if font in available_fonts:
                        plt.rcParams['font.sans-serif'] = [font, 'Times New Roman']
                        break
            
        except Exception as e:
            print(f"字体设置警告: {e}")
    
    def create_parameter_inputs(self):
        """为每个sheet创建横截面积输入框"""
        # 清除现有的输入框
        for widget in self.param_frame.winfo_children():
            widget.destroy()
        
        # 如果没有数据，显示提示
        if not self.excel_data:
            ttk.Label(self.param_frame, text="请先加载Excel数据", 
                     style="Medium.TLabel").pack()
            return
        
        # 创建标题
        title_frame = ttk.Frame(self.param_frame)
        title_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(title_frame, text="Sheet名称", width=20, 
                 style="Medium.TLabel").pack(side=tk.LEFT, padx=5)
        ttk.Label(title_frame, text="横截面积 (mm²)", width=15, 
                 style="Medium.TLabel").pack(side=tk.LEFT, padx=5)
        
        # 为每个sheet创建输入行
        self.area_entries = {}
        for i, sheet_name in enumerate(self.excel_data.keys()):
            row_frame = ttk.Frame(self.param_frame)
            row_frame.pack(fill=tk.X, pady=2)
            
            ttk.Label(row_frame, text=sheet_name, width=20, 
                     style="Small.TLabel").pack(side=tk.LEFT, padx=5)
            
            entry_var = tk.StringVar()
            # 如果有之前的值，恢复它
            if sheet_name in self.cross_sectional_areas:
                entry_var.set(str(self.cross_sectional_areas[sheet_name]))
            
            entry = ttk.Entry(row_frame, textvariable=entry_var, width=15,
                             style="Small.TEntry")
            entry.pack(side=tk.LEFT, padx=5)
            
            self.area_entries[sheet_name] = entry
        
        # 添加确认按钮
        ttk.Button(self.param_frame, text="确认所有参数", 
                  command=self.set_all_parameters,
                  style="Medium.TButton").pack(pady=(10, 0))
    
    def set_all_parameters(self):
        """设置所有sheet的参数"""
        success_count = 0
        error_sheets = []
        
        for sheet_name, entry in self.area_entries.items():
            try:
                area_text = entry.get().strip()
                if not area_text:
                    error_sheets.append(f"{sheet_name}: 未输入横截面积")
                    continue
                
                area = float(area_text)
                if area <= 0:
                    error_sheets.append(f"{sheet_name}: 横截面积必须大于0")
                    continue
                
                self.cross_sectional_areas[sheet_name] = area
                success_count += 1
                
            except ValueError:
                error_sheets.append(f"{sheet_name}: 无效的数值")
        
        if error_sheets:
            messagebox.showwarning("警告", 
                f"成功设置 {success_count} 个sheet的参数\n"
                f"以下sheet参数设置失败:\n" + "\n".join(error_sheets))
        else:
            messagebox.showinfo("成功", f"已成功设置 {success_count} 个sheet的参数")
        
        # 生成同名csv文件保存截面尺寸数据
        self.save_sectional_area_to_csv()
    
    def set_parameters(self):
        """设置测试参数"""
        try:
            area = float(self.area_entry.get())
            if area <= 0:
                messagebox.showerror("错误", "横截面积必须大于0")
                return
            self.cross_sectional_area = area
            messagebox.showinfo("成功", f"参数已设置：横截面积 = {area} mm²")
        except ValueError:
            messagebox.showerror("错误", "请输入有效的横截面积数值")
    
    def load_config(self):
        """加载配置文件"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    if 'cross_sectional_areas' in config:
                        self.cross_sectional_areas = config['cross_sectional_areas']
                    if 'legend_texts' in config:
                        self.legend_texts = config['legend_texts']
        except Exception as e:
            print(f"加载配置文件失败: {e}")
    
    def check_for_csv_config(self, excel_file_path):
        """检查同文件夹下是否存在同名csv文件，并加载截面尺寸数据"""
        try:
            # 获取Excel文件的目录和文件名
            excel_dir = os.path.dirname(excel_file_path)
            excel_filename = os.path.basename(excel_file_path)
            
            # 生成同名csv文件路径
            csv_filename = os.path.splitext(excel_filename)[0] + '.csv'
            csv_file_path = os.path.join(excel_dir, csv_filename)
            
            if os.path.exists(csv_file_path):
                # 读取csv文件
                df_config = pd.read_csv(csv_file_path)
                
                # 检查必要的列
                if 'sheet_name' in df_config.columns and 'cross_sectional_area' in df_config.columns:
                    # 加载截面尺寸数据
                    loaded_count = 0
                    for _, row in df_config.iterrows():
                        sheet_name = str(row['sheet_name'])
                        cross_sectional_area = float(row['cross_sectional_area'])
                        
                        # 只有当该sheet存在于当前加载的Excel文件中时，才使用这些数据
                        if sheet_name in self.excel_data:
                            self.cross_sectional_areas[sheet_name] = cross_sectional_area
                            loaded_count += 1
                    
                    if loaded_count > 0:
                        messagebox.showinfo("成功", f"已从 {csv_filename} 加载 {loaded_count} 个sheet的截面尺寸数据")
                
        except Exception as e:
            print(f"检查csv配置文件时出错: {e}")
    
    def save_sectional_area_to_csv(self):
        """将截面尺寸数据保存到与Excel同名的csv文件中"""
        try:
            if not self.current_excel_path or not self.excel_data:
                print("未加载Excel文件或没有数据，跳过保存")
                return
            
            # 获取Excel文件的目录和文件名
            excel_dir = os.path.dirname(self.current_excel_path)
            excel_filename = os.path.basename(self.current_excel_path)
            
            # 生成同名csv文件路径
            csv_filename = os.path.splitext(excel_filename)[0] + '.csv'
            csv_file_path = os.path.join(excel_dir, csv_filename)
            
            # 准备保存的数据
            config_data = []
            for sheet_name in self.excel_data.keys():
                if sheet_name in self.cross_sectional_areas:
                    config_data.append({
                        'sheet_name': sheet_name,
                        'cross_sectional_area': self.cross_sectional_areas[sheet_name]
                    })
            
            if not config_data:
                print("没有可用的截面尺寸数据，跳过保存")
                return
            
            # 创建DataFrame并保存为csv
            df_config = pd.DataFrame(config_data)
            df_config.to_csv(csv_file_path, index=False, encoding='utf-8')
            
            print(f"截面尺寸数据已保存到: {csv_file_path}")
            messagebox.showinfo("成功", f"截面尺寸数据已保存到 {csv_filename}")
            
        except Exception as e:
            print(f"保存截面尺寸数据到csv时出错: {e}")
            messagebox.showerror("错误", f"保存截面尺寸数据时出错: {e}")
    
    def save_config(self):
        """保存配置文件"""
        try:
            config = {
                'cross_sectional_areas': self.cross_sectional_areas,
                'legend_texts': self.legend_texts
            }
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存配置文件失败: {e}")
    
    def on_close(self):
        """窗口关闭事件处理"""
        self.save_config()
        
        # 释放matplotlib资源
        if hasattr(self, 'fig'):
            plt.close(self.fig)
        plt.close('all')
        plt.clf()  # 清除当前图形
        plt.cla()  # 清除当前轴
        plt.close()  # 关闭当前窗口
        
        # 销毁所有Tkinter窗口
        self.root.quit()  # 退出主循环
        self.root.destroy()  # 销毁窗口
        
        # 强制清理资源
        import gc
        gc.collect()
    
    def load_excel_data(self):
        """从Excel文件加载数据"""
        file_path = filedialog.askopenfilename(
            title="选择Excel数据文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        
        if file_path:
            # 保存当前Excel文件路径
            self.current_excel_path = file_path
            try:
                # 读取Excel文件的所有sheet名称
                excel_file = pd.ExcelFile(file_path)
                sheet_names = excel_file.sheet_names
                
                if not sheet_names:
                    messagebox.showerror("错误", "Excel文件中没有sheet")
                    return
                
                # 清空之前的数据
                self.excel_data.clear()
                
                # 读取每个sheet的数据
                for sheet_name in sheet_names:
                    try:
                        # 读取sheet数据
                        df = pd.read_excel(file_path, sheet_name=sheet_name)
                        
                        # 查找所需的列
                        load_col = None
                        extensometer_col = None
                        
                        # 查找载荷列（可能包含'载荷'或'Load'）
                        for col in df.columns:
                            if isinstance(col, str):
                                col_lower = col.lower()
                                if '载荷' in col or 'load' in col_lower or 'force' in col_lower:
                                    load_col = col
                                elif '引伸' in col or 'extenso' in col_lower or 'strain' in col_lower:
                                    extensometer_col = col
                        
                        # 如果没找到中文列名，尝试使用第一行数据作为列名
                        if load_col is None or extensometer_col is None:
                            # 使用第二行作为表头（假设第一行可能是单位）
                            df_alternative = pd.read_excel(file_path, sheet_name=sheet_name, header=1)
                            for col in df_alternative.columns:
                                if isinstance(col, str):
                                    col_lower = col.lower()
                                    if '载荷' in col or 'load' in col_lower or 'force' in col_lower:
                                        load_col = col
                                    elif '引伸' in col or 'extenso' in col_lower or 'strain' in col_lower:
                                        extensometer_col = col
                            
                            if load_col is not None and extensometer_col is not None:
                                df = df_alternative
                        
                        # 如果还是没找到，尝试基于位置（假设第1列是载荷，第3列是引伸计）
                        if load_col is None or extensometer_col is None:
                            if len(df.columns) >= 4:
                                # 尝试识别数据列
                                for i, col in enumerate(df.columns):
                                    if df[col].dtype in ['float64', 'int64']:
                                        if load_col is None:
                                            load_col = col
                                        elif extensometer_col is None:
                                            extensometer_col = col
                                            break
                        
                        if load_col is not None and extensometer_col is not None:
                            # 提取所需的两列数据
                            extracted_data = pd.DataFrame({
                                'Load_N': pd.to_numeric(df[load_col], errors='coerce'),
                                'Displacement_mm': pd.to_numeric(df[extensometer_col], errors='coerce')
                            })
                            
                            # 删除NaN值
                            extracted_data = extracted_data.dropna()
                            
                            # 确保数据量足够
                            if len(extracted_data) > 10:
                                # 存储数据
                                self.excel_data[sheet_name] = extracted_data
                                print(f"Sheet '{sheet_name}': 找到 {len(extracted_data)} 行数据")
                            else:
                                print(f"Sheet '{sheet_name}': 数据量不足，已跳过")
                        else:
                            print(f"Sheet '{sheet_name}': 未找到所需的列")
                        
                    except Exception as e:
                        print(f"读取sheet '{sheet_name}'时出错: {str(e)}")
                
                if not self.excel_data:
                    messagebox.showerror("错误", "未在任何sheet中找到所需的载荷和引伸计数据列")
                    return
                
                # 更新下拉框
                self.sheet_combobox['values'] = list(self.excel_data.keys())
                self.sheet_combobox.set(list(self.excel_data.keys())[0])
                
                # 检查同文件夹下是否存在同名csv文件
                self.check_for_csv_config(file_path)
                
                # 创建参数输入框
                self.create_parameter_inputs()
                
                # 初始化图例文本
                for sheet_name in self.excel_data.keys():
                    self.legend_texts[sheet_name] = sheet_name
                
                # 自动选择第一个sheet
                self.on_sheet_select(None)
                
                # 更新预览信息
                file_name = os.path.basename(file_path)
                self.preview_info_label.config(
                    text=f"已加载文件: {file_name}\n共 {len(self.excel_data)} 个sheet，总计 {sum(len(data) for data in self.excel_data.values())} 行数据"
                )
                
                messagebox.showinfo("成功", f"已成功加载 {len(self.excel_data)} 个sheet的数据")
                
            except Exception as e:
                messagebox.showerror("错误", f"读取Excel文件失败：{str(e)}")
    
    def on_sheet_select(self, event):
        """当选择不同的sheet时更新预览"""
        if self.sheet_combobox.get():
            sheet_name = self.sheet_combobox.get()
            self.current_sheet_name = sheet_name
            
            if sheet_name in self.excel_data:
                data = self.excel_data[sheet_name]
                
                # 更新预览文本框
                self.preview_text.delete(1.0, tk.END)
                
                # 显示前30行数据
                preview_lines = min(30, len(data))
                self.preview_text.insert(1.0, f"Sheet: {sheet_name}\n")
                self.preview_text.insert(tk.END, f"数据行数: {len(data)}\n")
                self.preview_text.insert(tk.END, "="*50 + "\n")
                self.preview_text.insert(tk.END, "载荷(N)              位移(mm)\n")
                self.preview_text.insert(tk.END, "-"*50 + "\n")
                
                for i in range(preview_lines):
                    self.preview_text.insert(tk.END, f"{data.iloc[i, 0]:>12.2f}        {data.iloc[i, 1]:>12.4f}\n")
                
                if len(data) > preview_lines:
                    self.preview_text.insert(tk.END, f"\n... 还有 {len(data) - preview_lines} 行数据\n")
    
    def calculate_yield_strength_robust(self, stress, strain):
        """更鲁棒的屈服强度计算方法 (0.2% 偏移法)"""
        if len(stress) < 20:
            return None, None
        
        try:
            # 方法1: 使用整体趋势，容忍局部波动
            # 对数据进行平滑处理
            # 增大移动平均窗口，使用更平滑的数据
            window_size = min(15, len(stress) // 8)
            if window_size < 5:
                window_size = 5
            
            # 使用Savitzky-Golay滤波器进行平滑，保留更多特征
            from scipy.signal import savgol_filter
            try:
                stress_smooth = savgol_filter(stress, window_length=window_size, polyorder=2)
                strain_smooth = strain
            except:
                # 如果Savitzky-Golay失败，回退到移动平均
                stress_smooth = np.convolve(stress, np.ones(window_size)/window_size, mode='valid')
                strain_smooth = strain[window_size-1:]
            
            # 方法2: 使用应力增量法确定弹性阶段
            # 寻找初始线性段（应力变化相对稳定的区域）
            strain_increments = np.diff(strain_smooth)
            stress_increments = np.diff(stress_smooth)
            
            # 计算应变-应力比（近似弹性模量）
            ratios = stress_increments / (strain_increments + 1e-10)
            
            # 寻找比值相对稳定的区域 - 使用更大的初始窗口
            initial_window = min(30, len(ratios) // 3)
            ratio_mean = np.mean(ratios[:initial_window])
            ratio_std = np.std(ratios[:initial_window])
            
            # 寻找弹性阶段的结束点 - 允许更大的波动，避免过早截断
            elastic_end = len(ratios)
            tolerance = 3.0  # 增加容忍度到3倍标准差
            
            # 滑动窗口检查弹性阶段
            sliding_window = min(10, len(ratios) // 20)
            if sliding_window < 3:
                sliding_window = 3
            
            for i in range(initial_window, len(ratios) - sliding_window):
                window_ratios = ratios[i:i+sliding_window]
                window_mean = np.mean(window_ratios)
                if abs(window_mean - ratio_mean) > tolerance * ratio_std:
                    elastic_end = i
                    break
            
            # 确保弹性阶段有足够的数据点
            if elastic_end < 10:
                elastic_end = min(30, len(stress_smooth) // 2)
            
            # 线性拟合弹性阶段
            x_elastic = strain_smooth[:elastic_end]
            y_elastic = stress_smooth[:elastic_end]
            
            if len(x_elastic) < 5:
                return None, None
            
            A = np.vstack([x_elastic, np.ones(len(x_elastic))]).T
            m, c = np.linalg.lstsq(A, y_elastic, rcond=None)[0]
            
            # 0.2% 塑性应变偏移线
            offset_strain = strain + 0.002
            offset_line = m * offset_strain + c
            
            # 寻找与偏移线的交点
            # 从弹性阶段结束点开始找，但使用原始数据点
            search_start = max(0, elastic_end - window_size + 1)
            
            for i in range(search_start, len(stress)-1):
                if stress[i] <= offset_line[i] and stress[i+1] > offset_line[i+1]:
                    # 线性插值找到精确交点
                    x1, x2 = strain[i], strain[i+1]
                    y1, y2 = stress[i] - offset_line[i], stress[i+1] - offset_line[i+1]
                    
                    if y1 * y2 < 0:  # 异号，说明有交点
                        t = -y1 / (y2 - y1)
                        yield_strain = x1 + t * (x2 - x1)
                        yield_strength = m * (yield_strain + 0.002) + c
                        return yield_strength, yield_strain
            
            # 如果没有找到交点，尝试使用更鲁棒的方法
            # 方法1: 使用更大范围的数据重新拟合
            # 使用前40%的数据进行拟合
            second_try_end = min(len(stress_smooth) // 2, 100)
            if second_try_end > elastic_end:
                x_elastic_2 = strain_smooth[:second_try_end]
                y_elastic_2 = stress_smooth[:second_try_end]
                
                if len(x_elastic_2) >= 5:
                    A2 = np.vstack([x_elastic_2, np.ones(len(x_elastic_2))]).T
                    m2, c2 = np.linalg.lstsq(A2, y_elastic_2, rcond=None)[0]
                    
                    offset_line_2 = m2 * (strain + 0.002) + c2
                    
                    # 重新寻找交点
                    for i in range(search_start, len(stress)-1):
                        if stress[i] <= offset_line_2[i] and stress[i+1] > offset_line_2[i+1]:
                            x1, x2 = strain[i], strain[i+1]
                            y1, y2 = stress[i] - offset_line_2[i], stress[i+1] - offset_line_2[i+1]
                            
                            if y1 * y2 < 0:
                                t = -y1 / (y2 - y1)
                                yield_strain = x1 + t * (x2 - x1)
                                yield_strength = m2 * (yield_strain + 0.002) + c2
                                return yield_strength, yield_strain
            
            # 方法2: 使用0.2%应变偏移法的标准实现
            # 寻找弹性模量的最佳估计
            # 计算整个曲线的弹性模量（使用初始线性部分）
            if len(stress) > 50:
                # 使用前20-30%的数据进行拟合
                fit_end = min(int(len(stress) * 0.3), 100)
                x_fit = strain[:fit_end]
                y_fit = stress[:fit_end]
                
                A3 = np.vstack([x_fit, np.ones(len(x_fit))]).T
                m3, c3 = np.linalg.lstsq(A3, y_fit, rcond=None)[0]
                
                # 创建偏移线
                offset_line_3 = m3 * (strain + 0.002) + c3
                
                # 寻找交点
                for i in range(0, len(stress)-1):
                    if stress[i] <= offset_line_3[i] and stress[i+1] > offset_line_3[i+1]:
                        x1, x2 = strain[i], strain[i+1]
                        y1, y2 = stress[i] - offset_line_3[i], stress[i+1] - offset_line_3[i+1]
                        
                        if y1 * y2 < 0:
                            t = -y1 / (y2 - y1)
                            yield_strain = x1 + t * (x2 - x1)
                            yield_strength = m3 * (yield_strain + 0.002) + c3
                            return yield_strength, yield_strain
            
            # 如果所有方法都失败，返回最大应力的90%作为近似
            if len(stress) > 0:
                max_stress = np.max(stress)
                max_strain = strain[np.argmax(stress)]
                return 0.9 * max_stress, max_strain * 0.9
            
            return None, None
            
        except Exception as e:
            print(f"屈服强度计算错误: {e}")
            # 如果scipy导入失败，尝试不使用它的版本
            try:
                # 简化版实现，不使用scipy
                window_size = min(15, len(stress) // 8)
                if window_size < 5:
                    window_size = 5
                
                # 移动平均
                stress_smooth = np.convolve(stress, np.ones(window_size)/window_size, mode='valid')
                strain_smooth = strain[window_size-1:]
                
                # 使用更大的初始窗口和容忍度
                initial_window = min(30, len(stress_smooth) // 3)
                x_elastic = strain_smooth[:initial_window]
                y_elastic = stress_smooth[:initial_window]
                
                if len(x_elastic) >= 5:
                    A = np.vstack([x_elastic, np.ones(len(x_elastic))]).T
                    m, c = np.linalg.lstsq(A, y_elastic, rcond=None)[0]
                    
                    offset_line = m * (strain + 0.002) + c
                    
                    for i in range(0, len(stress)-1):
                        if stress[i] <= offset_line[i] and stress[i+1] > offset_line[i+1]:
                            x1, x2 = strain[i], strain[i+1]
                            y1, y2 = stress[i] - offset_line[i], stress[i+1] - offset_line[i+1]
                            
                            if y1 * y2 < 0:
                                t = -y1 / (y2 - y1)
                                yield_strain = x1 + t * (x2 - x1)
                                yield_strength = m * (yield_strain + 0.002) + c
                                return yield_strength, yield_strain
                
                # 最后尝试：使用最大应力的85-90%作为近似
                if len(stress) > 0:
                    max_stress = np.max(stress)
                    return 0.88 * max_stress, strain[np.argmax(stress)] * 0.88
            except Exception as e2:
                print(f"简化版计算也失败: {e2}")
            
            return None, None
    
    def calculate_tensile_properties(self, data, sheet_name):
        """计算拉伸性能参数"""
        if data is None or len(data) < 20:
            return None, None, None, "数据量不足（至少需要20个数据点）"
        
        load = data['Load_N'].values
        displacement = data['Displacement_mm'].values
        
        # 检查数据有效性
        if len(load) == 0 or len(displacement) == 0:
            return None, None, None, "数据为空"
        
        # 检查是否设置了该sheet的横截面积
        if sheet_name not in self.cross_sectional_areas:
            return None, None, None, f"未设置Sheet '{sheet_name}'的横截面积"
        
        cross_sectional_area = self.cross_sectional_areas[sheet_name]
        
        try:
            # 计算工程应力和工程应变
            stress = load / cross_sectional_area  # MPa
            strain = displacement / self.gauge_length
            
            # 抗拉强度（最大应力）
            tensile_strength = np.max(stress)
            
            # 屈服强度（使用鲁棒的方法）
            yield_strength, yield_strain = self.calculate_yield_strength_robust(stress, strain)
            
            # 延伸率（最大应变对应的延伸率）
            max_strain = np.max(strain)
            elongation = max_strain * 100  # 转换为百分比
            
            error_msg = ""
            if yield_strength is None:
                error_msg = "屈服强度计算失败"
            
            return yield_strength, tensile_strength, elongation, error_msg
            
        except Exception as e:
            return None, None, None, f"计算错误: {str(e)}"
    
    def process_current_sheet(self):
        """处理当前选中的sheet数据"""
        if not self.current_sheet_name or self.current_sheet_name not in self.excel_data:
            messagebox.showerror("错误", "请先加载Excel数据并选择sheet")
            return
        
        # 获取当前sheet数据
        data = self.excel_data[self.current_sheet_name]
        
        # 计算性能参数
        yield_strength, tensile_strength, elongation, error_msg = self.calculate_tensile_properties(data, self.current_sheet_name)
        
        # 显示结果
        self.results_text.config(state='normal')
        self.results_text.delete(1.0, tk.END)
        
        results = f"当前Sheet: {self.current_sheet_name}\n"
        results += f"数据点数: {len(data)}\n"
        results += f"横截面积: {self.cross_sectional_areas.get(self.current_sheet_name, '未设置')} mm²\n"
        results += "="*40 + "\n"
        results += "计算结果：\n\n"
        
        if yield_strength:
            results += f"屈服强度 (Rp0.2): {yield_strength:.2f} MPa\n"
        else:
            results += f"屈服强度: 计算失败\n"
            if error_msg:
                results += f"原因: {error_msg}\n"
        
        if tensile_strength:
            results += f"抗拉强度 (Rm): {tensile_strength:.2f} MPa\n"
        
        if elongation:
            results += f"延伸率 (A): {elongation:.2f} %\n"
        
        self.results_text.insert(1.0, results)
        self.results_text.config(state='disabled')
        
        # 绘制曲线
        self.plot_sheet_data(data, self.current_sheet_name)
    
    def process_all_sheets(self):
        """批量处理所有sheet数据"""
        if not self.excel_data:
            messagebox.showerror("错误", "请先加载Excel数据")
            return
        
        # 检查是否所有sheet都设置了横截面积
        sheets_without_area = [name for name in self.excel_data.keys() 
                              if name not in self.cross_sectional_areas]
        if sheets_without_area:
            messagebox.showwarning("警告", 
                f"以下sheet未设置横截面积:\n" + "\n".join(sheets_without_area) + 
                "\n\n将跳过这些sheet的计算。")
        
        # 处理所有sheet
        all_results = []
        
        for sheet_name, data in self.excel_data.items():
            if sheet_name not in self.cross_sectional_areas:
                continue
            
            yield_strength, tensile_strength, elongation, error_msg = self.calculate_tensile_properties(data, sheet_name)
            
            all_results.append({
                'sheet_name': sheet_name,
                'data_points': len(data),
                'cross_sectional_area': self.cross_sectional_areas[sheet_name],
                'yield_strength': yield_strength,
                'tensile_strength': tensile_strength,
                'elongation': elongation,
                'error_msg': error_msg
            })
        
        if not all_results:
            messagebox.showerror("错误", "没有可以计算的sheet")
            return
        
        # 显示多sheet结果
        self.multi_results_text.config(state='normal')
        self.multi_results_text.delete(1.0, tk.END)
        
        results_text = "所有Sheet计算结果汇总:\n"
        results_text += "="*60 + "\n\n"
        
        for result in all_results:
            results_text += f"Sheet: {result['sheet_name']}\n"
            results_text += f"数据点数: {result['data_points']}\n"
            results_text += f"横截面积: {result['cross_sectional_area']} mm²\n"
            
            if result['yield_strength']:
                results_text += f"屈服强度: {result['yield_strength']:.2f} MPa\n"
            else:
                results_text += "屈服强度: 计算失败\n"
                if result['error_msg']:
                    results_text += f"原因: {result['error_msg']}\n"
            
            if result['tensile_strength']:
                results_text += f"抗拉强度: {result['tensile_strength']:.2f} MPa\n"
            
            if result['elongation']:
                results_text += f"延伸率: {result['elongation']:.2f} %\n"
            
            results_text += "-"*40 + "\n\n"
        
        self.multi_results_text.insert(1.0, results_text)
        self.multi_results_text.config(state='disabled')
        
        # 绘制所有sheet的曲线对比
        self.plot_all_sheets()
        
        messagebox.showinfo("完成", f"已处理 {len(all_results)} 个sheet的数据")
    
    def plot_sheet_data(self, data, sheet_name):
        """绘制单个sheet的载荷-位移曲线"""
        if data is None or len(data) < 2:
            return
        
        self.ax.clear()
        
        load = data['Load_N'].values
        displacement = data['Displacement_mm'].values
        
        # 计算应力和应变
        if sheet_name in self.cross_sectional_areas:
            cross_sectional_area = self.cross_sectional_areas[sheet_name]
            stress = load / cross_sectional_area
            strain = displacement / self.gauge_length
            
            # 绘制应力-应变曲线
            legend_text = self.legend_texts.get(sheet_name, sheet_name)
            self.ax.plot(strain, stress, 'b-', linewidth=2.5, label=legend_text)
            
            # 标记关键点
            max_stress_idx = np.argmax(stress)
            if max_stress_idx < len(strain):
                self.ax.plot(strain[max_stress_idx], stress[max_stress_idx], 'ro', 
                           markersize=10, label=f'抗拉强度: {stress[max_stress_idx]:.1f} MPa')
            
            # 计算并标记屈服点（确保与计算结果一致）
            yield_strength, yield_strain = self.calculate_yield_strength_robust(stress, strain)
            if yield_strength and yield_strain:
                # 确保屈服点在曲线上，找到最接近的点（同时考虑应变和应力）
                # 首先找到最大应力点，屈服点应该在最大应力点之前
                max_stress_idx = np.argmax(stress)
                
                # 只在最大应力点之前的区域搜索屈服点
                search_region_strain = strain[:max_stress_idx+1]
                search_region_stress = stress[:max_stress_idx+1]
                
                # 计算搜索区域内每个数据点到计算点的欧几里得距离
                distances = np.sqrt((search_region_strain - yield_strain)**2 + (search_region_stress - yield_strength)**2)
                closest_idx_in_region = np.argmin(distances)
                closest_idx = closest_idx_in_region  # 转换回原始数组的索引
                
                closest_yield_strain = strain[closest_idx]
                closest_yield_strength = stress[closest_idx]
                
                # 绘制实际曲线上的屈服点标记
                self.ax.plot(closest_yield_strain, closest_yield_strength, 'go', markersize=10,
                           label=f'屈服强度: {yield_strength:.1f} MPa')
                # 可以选择添加一条虚线连接计算点和实际点
                # self.ax.plot([yield_strain, closest_yield_strain], [yield_strength, closest_yield_strength], 'g--', alpha=0.5)
            
            # 设置图形属性 - 去除标题
            self.ax.set_xlabel('应变', fontsize=14)
            self.ax.set_ylabel('应力 (MPa)', fontsize=14)
            
            # 设置刻度字体
            self.ax.tick_params(axis='both', which='major', labelsize=12)
            
            self.ax.grid(True, alpha=0.3, linestyle='--')
            self.ax.legend(loc='best', fontsize=12)
        else:
            # 直接绘制载荷-位移曲线
            legend_text = self.legend_texts.get(sheet_name, sheet_name)
            self.ax.plot(displacement, load, 'b-', linewidth=2.5, label=legend_text)
            self.ax.set_xlabel('位移 (mm)', fontsize=14)
            self.ax.set_ylabel('载荷 (N)', fontsize=14)
            
            # 设置刻度字体
            self.ax.tick_params(axis='both', which='major', labelsize=12)
            
            self.ax.grid(True, alpha=0.3, linestyle='--')
            self.ax.legend(loc='best', fontsize=12)
        
        self.fig.tight_layout()
        self.canvas.draw()
    
    def plot_all_sheets(self):
        """绘制所有sheet的曲线对比"""
        if not self.excel_data:
            return
        
        self.ax.clear()
        
        # 定义颜色和线型
        colors = ['blue', 'green', 'red', 'cyan', 'magenta', 'orange', 'purple', 'brown']
        linestyles = ['-', '--', '-.', ':']
        
        # 绘制每个sheet的曲线
        for i, (sheet_name, data) in enumerate(self.excel_data.items()):
            if len(data) < 10:
                continue
            
            if sheet_name not in self.cross_sectional_areas:
                continue
            
            load = data['Load_N'].values
            displacement = data['Displacement_mm'].values
            
            cross_sectional_area = self.cross_sectional_areas[sheet_name]
            stress = load / cross_sectional_area
            strain = displacement / self.gauge_length
            
            color = colors[i % len(colors)]
            linestyle = linestyles[(i // len(colors)) % len(linestyles)]
            
            # 使用自定义的图例文本
            legend_text = self.legend_texts.get(sheet_name, sheet_name)
            
            self.ax.plot(strain, stress, color=color, linestyle=linestyle, 
                       linewidth=2, label=legend_text, alpha=0.8)
        
        # 设置图形属性 - 去除标题
        self.ax.set_xlabel('应变', fontsize=14)
        self.ax.set_ylabel('应力 (MPa)', fontsize=14)
        
        # 设置刻度字体
        self.ax.tick_params(axis='both', which='major', labelsize=12)
        
        self.ax.grid(True, alpha=0.3, linestyle='--')
        self.ax.legend(loc='best', fontsize=11)
        
        self.fig.tight_layout()
        self.canvas.draw()
    
    def edit_legend_texts(self):
        """编辑图例文本"""
        if not self.excel_data:
            messagebox.showinfo("提示", "请先加载数据")
            return
        
        # 创建编辑窗口
        edit_window = tk.Toplevel(self.root)
        edit_window.title("编辑图例文本")
        edit_window.geometry("500x400")
        
        # 创建滚动框架
        canvas = tk.Canvas(edit_window)
        scrollbar = ttk.Scrollbar(edit_window, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 为每个sheet创建编辑框
        self.legend_entries = {}
        
        for i, sheet_name in enumerate(self.excel_data.keys()):
            row_frame = ttk.Frame(scrollable_frame)
            row_frame.pack(fill=tk.X, pady=5, padx=10)
            
            ttk.Label(row_frame, text=sheet_name, width=30).pack(side=tk.LEFT)
            
            entry_var = tk.StringVar(value=self.legend_texts.get(sheet_name, sheet_name))
            entry = ttk.Entry(row_frame, textvariable=entry_var, width=30)
            entry.pack(side=tk.LEFT, padx=10)
            
            self.legend_entries[sheet_name] = entry_var
        
        # 添加按钮
        button_frame = ttk.Frame(edit_window)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="确认", 
                  command=lambda: self.save_legend_texts(edit_window)).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="取消", 
                  command=edit_window.destroy).pack(side=tk.LEFT, padx=5)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
    def save_legend_texts(self, edit_window):
        """保存图例文本"""
        for sheet_name, entry_var in self.legend_entries.items():
            new_text = entry_var.get().strip()
            if new_text:
                self.legend_texts[sheet_name] = new_text
        
        edit_window.destroy()
        
        # 重新绘制图形
        if self.current_sheet_name:
            self.plot_sheet_data(self.excel_data[self.current_sheet_name], self.current_sheet_name)
        else:
            self.plot_all_sheets()
    
    def reset_legend_texts(self):
        """重置图例文本为sheet名称"""
        for sheet_name in self.excel_data.keys():
            self.legend_texts[sheet_name] = sheet_name
        
        # 重新绘制图形
        if self.current_sheet_name:
            self.plot_sheet_data(self.excel_data[self.current_sheet_name], self.current_sheet_name)
        else:
            self.plot_all_sheets()
    
    def save_plot(self):
        """保存图表"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG文件", "*.png"), ("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        
        if file_path:
            # 保存时重新应用字体设置
            try:
                # 备份当前字体设置
                original_font = rcParams['font.sans-serif']
                
                # 强制使用特定字体
                rcParams['font.sans-serif'] = ['SimSun', 'DejaVu Sans']
                
                # 设置保存图形的DPI和质量
                dpi = 1200
                if file_path.endswith('.png'):
                    self.fig.savefig(file_path, dpi=dpi, bbox_inches='tight', 
                                   facecolor='white', edgecolor='none')
                else:
                    self.fig.savefig(file_path, dpi=dpi, bbox_inches='tight', 
                                   facecolor='white', edgecolor='none')
                
                # 恢复字体设置
                rcParams['font.sans-serif'] = original_font
                
                messagebox.showinfo("成功", f"图表已保存到：{file_path}")
                
            except Exception as e:
                messagebox.showerror("错误", f"保存图表失败：{str(e)}")
    
    def export_all_results(self):
        """导出所有结果到文件"""
        if not self.excel_data:
            messagebox.showerror("错误", "没有可导出的结果")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("CSV文件", "*.csv"), ("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        
        if file_path:
            try:
                # 收集所有结果
                all_results = []
                
                for sheet_name, data in self.excel_data.items():
                    if sheet_name not in self.cross_sectional_areas:
                        continue
                    
                    yield_strength, tensile_strength, elongation, error_msg = self.calculate_tensile_properties(data, sheet_name)
                    
                    all_results.append({
                        'Sheet名称': sheet_name,
                        '数据点数': len(data),
                        '横截面积_mm²': self.cross_sectional_areas[sheet_name],
                        '屈服强度_MPa': round(yield_strength, 2) if yield_strength else '',
                        '抗拉强度_MPa': round(tensile_strength, 2) if tensile_strength else '',
                        '延伸率_%': round(elongation, 2) if elongation else '',
                        '备注': error_msg if error_msg else '计算成功'
                    })
                
                if not all_results:
                    messagebox.showerror("错误", "没有可以导出的结果")
                    return
                
                # 创建DataFrame
                results_df = pd.DataFrame(all_results)
                
                # 根据文件类型保存
                if file_path.endswith('.csv'):
                    results_df.to_csv(file_path, index=False, encoding='utf-8-sig')
                elif file_path.endswith('.xlsx'):
                    # 创建Excel写入器
                    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                        results_df.to_excel(writer, sheet_name='计算结果汇总', index=False)
                        
                        # 也可以保存原始数据
                        for sheet_name, data in self.excel_data.items():
                            if sheet_name not in self.cross_sectional_areas:
                                continue
                            # 只保存前1000行原始数据
                            data_to_save = data.head(1000)
                            data_to_save.to_excel(writer, sheet_name=f'{sheet_name[:30]}_原始数据', index=False)
                else:
                    # 保存为文本文件
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write("铍镍铜拉伸测试结果汇总\n")
                        f.write("="*70 + "\n\n")
                        f.write(f"引伸计标距: {self.gauge_length} mm\n")
                        f.write(f"测试样本数: {len(all_results)}\n\n")
                        
                        for result in all_results:
                            f.write(f"Sheet: {result['Sheet名称']}\n")
                            f.write(f"数据点数: {result['数据点数']}\n")
                            f.write(f"横截面积: {result['横截面积_mm²']} mm²\n")
                            
                            if result['屈服强度_MPa']:
                                f.write(f"屈服强度: {result['屈服强度_MPa']:.2f} MPa\n")
                            else:
                                f.write("屈服强度: N/A\n")
                            
                            f.write(f"抗拉强度: {result['抗拉强度_MPa']:.2f} MPa\n")
                            f.write(f"延伸率: {result['延伸率_%']:.2f} %\n")
                            f.write(f"备注: {result['备注']}\n")
                            f.write("-"*50 + "\n\n")
                
                messagebox.showinfo("成功", f"结果已导出到：{file_path}")
                
            except Exception as e:
                messagebox.showerror("错误", f"导出失败：{str(e)}")

def main():
    try:
        root = tk.Tk()
        
        # 设置窗口图标（可选）
        try:
            root.iconbitmap(default='')  # 可以设置图标路径
        except Exception as e:
            print(f"设置图标失败: {e}")
        
        app = TensileTestAnalyzer(root)
        root.mainloop()
    except Exception as e:
        import traceback
        print(f"程序错误: {e}")
        print("错误堆栈:")
        traceback.print_exc()
        input("按任意键退出...")

if __name__ == "__main__":
    main()