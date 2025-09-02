#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GPA计算器 - 图形界面版
使用tkinter创建用户友好的图形界面
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import os
from typing import Optional, Dict, Any

class GPACalculatorGUI:
    """GPA计算器图形界面类"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("GPA计算器")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        # 数据变量
        self.file_path = ""
        self.gpa_result = None
        
        # 创建界面
        self.setup_ui()
        
        # 设置样式
        self.setup_styles()
    
    def get_chinese_font(self, size: int = 12, weight: str = "normal") -> tuple:
        """获取支持中文的字体，带回退机制"""
        # 字体优先级列表（从最佳到备选）
        fonts = [
            "Noto Sans CJK SC",      # Google Noto字体
            "WenQuanYi Zen Hei",     # 文泉驿正黑
            "WenQuanYi Micro Hei",   # 文泉驿微米黑
            "DejaVu Sans",           # DejaVu字体
            "Liberation Sans",       # Liberation字体
            "sans-serif"             # 系统默认无衬线字体
        ]
        
        # 测试字体可用性
        import tkinter.font as tkfont
        for font_name in fonts:
            try:
                test_font = tkfont.Font(family=font_name, size=size, weight=weight)
                # 简单测试字体是否可用
                test_font.measure("测试")
                return (font_name, size, weight)
            except:
                continue
        
        # 如果所有字体都不可用，使用默认字体
        return ("TkDefaultFont", size, weight)

    def setup_styles(self):
        """设置界面样式"""
        # 配置ttk样式
        style = ttk.Style()
        style.theme_use('clam')
        
        # 配置按钮样式 - 使用支持中文的字体
        title_font = self.get_chinese_font(16, "bold")
        button_font = self.get_chinese_font(12)
        result_font = self.get_chinese_font(12, "bold")
        
        style.configure("Title.TLabel", font=title_font)
        style.configure("Big.TButton", font=button_font, padding=10)
        style.configure("Result.TLabel", font=result_font)
    

    
    def select_file(self):
        """选择Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择成绩Excel文件",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialdir=os.getcwd()
        )
        
        if file_path:
            self.file_path = file_path
            filename = os.path.basename(file_path)
            self.file_label.config(text=f"已选择: {filename}", foreground="green")
            self.calculate_button.config(state="normal")
            
            # 清空之前的结果
            self.clear_results()
    
    def validate_excel_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """验证和清理Excel数据"""
        # 检查必需的列
        if '学分' not in df.columns:
            raise ValueError("未找到'学分'列，请检查Excel文件格式")
        
        if '绩点' not in df.columns:
            raise ValueError("未找到'绩点'列，请检查Excel文件格式")
        
        # 选择需要的列
        columns_to_use = ['学分', '绩点']
        if '课程名' in df.columns:
            columns_to_use.insert(0, '课程名')
        elif '课程名称' in df.columns:
            columns_to_use.insert(0, '课程名称')
        
        # 创建工作数据
        work_df = df[columns_to_use].copy()
        
        # 删除空值行
        initial_count = len(work_df)
        work_df = work_df.dropna(subset=['学分', '绩点'])
        removed_count = initial_count - len(work_df)
        
        if removed_count > 0:
            messagebox.showinfo("数据清理", f"已忽略 {removed_count} 行空数据")
        
        if len(work_df) == 0:
            raise ValueError("没有找到有效的数据行")
        
        # 转换为数字类型
        try:
            work_df['学分'] = pd.to_numeric(work_df['学分'])
            work_df['绩点'] = pd.to_numeric(work_df['绩点'])
        except Exception as e:
            raise ValueError(f"数据格式错误，学分和绩点必须为数字: {str(e)}")
        
        # 过滤无效数据
        work_df = work_df[(work_df['学分'] > 0) & (work_df['绩点'] >= 0)]
        
        if len(work_df) == 0:
            raise ValueError("没有有效的学分和绩点数据")
        
        return work_df
    
    def calculate_gpa(self):
        """计算GPA并显示结果"""
        if not self.file_path:
            messagebox.showerror("错误", "请先选择Excel文件")
            return
        
        try:
            # 显示计算进度
            self.gpa_label.config(text="正在计算中...", foreground="orange")
            self.root.update()
            
            # 读取Excel文件
            df = pd.read_excel(self.file_path)
            
            # 验证数据
            work_df = self.validate_excel_data(df)
            
            # 计算权重分数
            work_df['权重分数'] = work_df['学分'] * work_df['绩点']
            
            # 计算GPA
            total_credits = work_df['学分'].sum()
            total_weighted = work_df['权重分数'].sum()
            gpa = total_weighted / total_credits
            
            # 保存结果
            self.gpa_result = {
                'data': work_df,
                'gpa': gpa,
                'total_credits': total_credits,
                'total_weighted': total_weighted,
                'course_count': len(work_df)
            }
            
            # 显示结果
            self.display_results()
            
        except Exception as e:
            self.gpa_label.config(text="计算失败", foreground="red")
            messagebox.showerror("计算错误", str(e))
    
    def display_results(self):
        """显示计算结果"""
        if not self.gpa_result:
            return
        
        result = self.gpa_result
        gpa = result['gpa']
        
        # 更新GPA标签
        gpa_text = f"🎯 您的GPA: {gpa:.4f}"
        if gpa >= 3.5:
            color = "green"
            gpa_text += " (优秀!)"
        elif gpa >= 3.0:
            color = "blue"
            gpa_text += " (良好)"
        elif gpa >= 2.5:
            color = "orange"
            gpa_text += " (及格)"
        else:
            color = "red"
            gpa_text += " (需努力)"
        
        self.gpa_label.config(text=gpa_text, foreground=color)
        
        # 清空文本框
        self.result_text.delete(1.0, tk.END)
        
        # 格式化详细结果
        output = []
        output.append("=" * 80)
        output.append("                           GPA 计算详情")
        output.append("=" * 80)
        output.append("")
        
        # 表头
        data = result['data']
        if '课程名' in data.columns:
            output.append(f"{'课程名称':<30} {'学分':<8} {'绩点':<8} {'权重分数':<10}")
        elif '课程名称' in data.columns:
            output.append(f"{'课程名称':<30} {'学分':<8} {'绩点':<8} {'权重分数':<10}")
        else:
            output.append(f"{'序号':<8} {'学分':<8} {'绩点':<8} {'权重分数':<10}")
        
        output.append("-" * 70)
        
        # 课程详情
        for i, (_, row) in enumerate(data.iterrows()):
            if '课程名' in data.columns:
                course_name = str(row['课程名'])[:25]
                output.append(f"{course_name:<30} {row['学分']:<8.1f} {row['绩点']:<8.2f} {row['权重分数']:<10.2f}")
            elif '课程名称' in data.columns:
                course_name = str(row['课程名称'])[:25]
                output.append(f"{course_name:<30} {row['学分']:<8.1f} {row['绩点']:<8.2f} {row['权重分数']:<10.2f}")
            else:
                output.append(f"课程{i+1:<5} {row['学分']:<8.1f} {row['绩点']:<8.2f} {row['权重分数']:<10.2f}")
        
        output.append("")
        output.append("=" * 80)
        output.append(f"📚 课程总数: {result['course_count']} 门")
        output.append(f"📊 总学分: {result['total_credits']:.1f}")
        output.append(f"📈 总权重分数: {result['total_weighted']:.2f}")
        output.append(f"🎯 平均学分绩点(GPA): {result['gpa']:.4f}")
        output.append("=" * 80)
        
        # 添加成绩分析
        output.append("")
        output.append("📋 成绩分析:")
        excellent_courses = len(data[data['绩点'] >= 4.0])
        good_courses = len(data[(data['绩点'] >= 3.0) & (data['绩点'] < 4.0)])
        average_courses = len(data[(data['绩点'] >= 2.0) & (data['绩点'] < 3.0)])
        poor_courses = len(data[data['绩点'] < 2.0])
        
        output.append(f"  优秀 (绩点≥4.0): {excellent_courses} 门")
        output.append(f"  良好 (3.0≤绩点<4.0): {good_courses} 门")
        output.append(f"  一般 (2.0≤绩点<3.0): {average_courses} 门")
        output.append(f"  待提升 (绩点<2.0): {poor_courses} 门")
        
        # 显示结果
        result_text = "\n".join(output)
        self.result_text.insert(tk.END, result_text)
    
    def clear_results(self):
        """清空显示结果"""
        self.result_text.delete(1.0, tk.END)
        self.gpa_label.config(text="等待计算...", foreground="black")
        self.gpa_result = None
    
    def save_results(self):
        """保存结果到文件"""
        if not self.gpa_result:
            messagebox.showwarning("警告", "没有计算结果可保存")
            return
        
        save_path = filedialog.asksaveasfilename(
            title="保存计算结果",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if save_path:
            try:
                content = self.result_text.get(1.0, tk.END)
                with open(save_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                messagebox.showinfo("成功", f"结果已保存到: {save_path}")
            except Exception as e:
                messagebox.showerror("保存错误", f"保存文件时出错: {str(e)}")
    
    def setup_ui(self):
        """创建用户界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="🎓 GPA计算器", style="Title.TLabel")
        title_label.grid(row=0, column=0, pady=(0, 20))
        
        # 说明文字
        info_font = self.get_chinese_font(10)
        info_label = ttk.Label(
            main_frame, 
            text="上传包含学分和绩点的Excel文件，自动计算您的GPA",
            font=info_font,
            foreground="gray"
        )
        info_label.grid(row=1, column=0, pady=(0, 20))
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="📁 选择成绩文件", padding="15")
        file_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        file_frame.columnconfigure(1, weight=1)
        
        # 文件选择按钮
        self.select_button = ttk.Button(
            file_frame, 
            text="📂 浏览文件...", 
            command=self.select_file,
            style="Big.TButton"
        )
        self.select_button.grid(row=0, column=0, padx=(0, 15))
        
        # 文件路径显示
        self.file_label = ttk.Label(file_frame, text="请选择Excel文件 (.xlsx)", foreground="gray")
        self.file_label.grid(row=0, column=1, sticky=(tk.W, tk.E))
        
        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, pady=(0, 20))
        
        self.calculate_button = ttk.Button(
            button_frame,
            text="🧮 计算GPA",
            command=self.calculate_gpa,
            style="Big.TButton",
            state="disabled"
        )
        self.calculate_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.clear_button = ttk.Button(
            button_frame,
            text="🗑️ 清空结果",
            command=self.clear_results,
            style="Big.TButton"
        )
        self.clear_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.save_button = ttk.Button(
            button_frame,
            text="💾 保存结果",
            command=self.save_results,
            style="Big.TButton"
        )
        self.save_button.pack(side=tk.LEFT)
        
        # 结果显示区域
        result_frame = ttk.LabelFrame(main_frame, text="📊 计算结果", padding="15")
        result_frame.grid(row=4, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(1, weight=1)
        
        # GPA总结果显示
        self.gpa_label = ttk.Label(result_frame, text="等待计算...", style="Result.TLabel")
        self.gpa_label.grid(row=0, column=0, pady=(0, 15))
        
        # 详细结果显示
        text_font = self.get_chinese_font(10)
        self.result_text = scrolledtext.ScrolledText(
            result_frame,
            width=80,
            height=15,
            font=text_font,
            wrap=tk.WORD
        )
        self.result_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
    
    def run(self):
        """启动GUI"""
        # 设置窗口居中
        self.center_window()
        
        # 显示欢迎信息
        welcome_text = """欢迎使用GPA计算器！

使用说明：
1. 点击"浏览文件"选择您的成绩Excel文件
2. 确保文件包含"学分"和"绩点"列
3. 点击"计算GPA"获得结果
4. 可以保存计算结果到文本文件

支持的Excel格式：
• 学分列：学分、学时、credit等
• 绩点列：绩点、gpa、grade等  
• 课程名：课程名、课程名称（可选）

开始计算您的GPA吧！📚
"""
        self.result_text.insert(tk.END, welcome_text)
        
        # 启动主循环
        self.root.mainloop()
    
    def center_window(self):
        """将窗口居中显示"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

def main():
    """主函数"""
    app = GPACalculatorGUI()
    app.run()

if __name__ == "__main__":
    main()
