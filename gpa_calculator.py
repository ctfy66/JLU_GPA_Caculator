#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GPA计算器
功能：读取Excel文件中的课程学分和绩点，计算总GPA
公式：GPA = sum(学分 × 绩点) / sum(学分)
"""

import pandas as pd
import argparse
import sys
import os
from typing import Tuple, Dict, Any

class GPACalculator:
    """GPA计算器类"""
    
    def __init__(self):
        self.courses_data = None
        self.total_credits = 0
        self.weighted_points = 0
        self.gpa = 0
    
    def read_excel_file(self, file_path: str) -> pd.DataFrame:
        """
        读取Excel文件
        
        Args:
            file_path: Excel文件路径
            
        Returns:
            DataFrame: 包含课程数据的DataFrame
            
        Raises:
            FileNotFoundError: 文件不存在
            ValueError: 文件格式错误
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")
        
        if not file_path.endswith('.xlsx'):
            raise ValueError("请提供.xlsx格式的Excel文件")
        
        try:
            # 尝试读取Excel文件
            df = pd.read_excel(file_path)
            print(f"成功读取文件: {file_path}")
            print(f"文件包含 {len(df)} 行数据")
            return df
        except Exception as e:
            raise ValueError(f"读取Excel文件时出错: {str(e)}")
    
    def validate_data_format(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        验证并标准化数据格式
        
        Args:
            df: 原始DataFrame
            
        Returns:
            DataFrame: 清理后的DataFrame
            
        Raises:
            ValueError: 数据格式不正确
        """
        # 检查必需的列
        required_columns = ['学分', '绩点']
        column_mapping = {}
        
        # 尝试匹配列名（支持不同的命名方式）
        df_columns = df.columns.tolist()
        
        # 学分列的可能名称
        credit_names = ['学分', '学时', 'credit', 'credits', '学分数']
        grade_names = ['绩点', '成绩', 'gpa', 'grade', '绩点成绩']
        
        # 查找学分列
        credit_col = None
        for col in df_columns:
            if any(name in str(col).lower() for name in [name.lower() for name in credit_names]):
                credit_col = col
                break
        
        # 查找绩点列 - 精确匹配优先，避免误匹配
        grade_col = None
        # 首先尝试精确匹配
        for col in df_columns:
            if str(col).strip() == '绩点':
                grade_col = col
                break
        
        # 如果精确匹配失败，再尝试模糊匹配
        if grade_col is None:
            for col in df_columns:
                col_lower = str(col).lower()
                if col_lower == 'gpa' or col_lower == 'grade' or '绩点' in col_lower:
                    grade_col = col
                    break
        
        if credit_col is None:
            raise ValueError(f"未找到学分列。请确保Excel文件包含以下列名之一: {credit_names}")
        
        if grade_col is None:
            raise ValueError(f"未找到绩点列。请确保Excel文件包含以下列名之一: {grade_names}")
        
        # 重命名列为标准名称
        column_mapping[credit_col] = '学分'
        column_mapping[grade_col] = '绩点'
        df = df.rename(columns=column_mapping)
        
        # 检查课程名称列
        course_col = None
        course_names = ['课程', '课程名称', 'course', '科目', '课程名']
        for col in df_columns:
            for name in course_names:
                if str(col).strip() == name or name in str(col).lower():
                    course_col = col
                    break
            if course_col:
                break
        
        if course_col and course_col not in column_mapping.values():
            df = df.rename(columns={course_col: '课程名称'})
        
        # 过滤出有效数据，确保使用标准化后的列名
        available_columns = ['学分', '绩点']
        if '课程名称' in df.columns:
            available_columns.insert(0, '课程名称')
        elif '课程名' in df.columns:
            df = df.rename(columns={'课程名': '课程名称'})
            available_columns.insert(0, '课程名称')
        
        # 选择需要的列
        df_clean = df[available_columns].copy()
        
        # 删除空行
        df_clean = df_clean.dropna(subset=['学分', '绩点'])
        
        # 验证数据类型
        try:
            # 确保学分和绩点列存在且为Series类型
            if '学分' not in df_clean.columns or '绩点' not in df_clean.columns:
                raise ValueError("缺少必需的学分或绩点列")
            
            # 转换为数字，无法转换的设为NaN
            df_clean.loc[:, '学分'] = pd.to_numeric(df_clean['学分'], errors='coerce')
            df_clean.loc[:, '绩点'] = pd.to_numeric(df_clean['绩点'], errors='coerce')
            
        except Exception as e:
            raise ValueError(f"学分或绩点数据格式错误，请确保为数字: {str(e)}")
        
        # 删除无法转换为数字的行
        initial_count = len(df_clean)
        df_clean = df_clean.dropna(subset=['学分', '绩点'])
        removed_count = initial_count - len(df_clean)
        
        if removed_count > 0:
            print(f"已忽略 {removed_count} 行无效数据（学分或绩点为空/非数字）")
        
        if len(df_clean) == 0:
            raise ValueError("没有找到有效的学分和绩点数据")
        
        # 验证数据合理性
        if (df_clean['学分'] <= 0).any():
            raise ValueError("学分必须大于0")
        
        if (df_clean['绩点'] < 0).any() or (df_clean['绩点'] > 5).any():
            print("警告: 发现绩点超出常规范围(0-5)，请检查数据是否正确")
        
        return df_clean
    
    def calculate_gpa(self, df: pd.DataFrame) -> Dict[str, Any]:
        """
        计算GPA
        
        Args:
            df: 包含学分和绩点的DataFrame
            
        Returns:
            Dict: 包含计算结果的字典
        """
        # 计算每门课的权重分数
        df['权重分数'] = df['学分'] * df['绩点']
        
        # 计算总学分和总权重分数
        total_credits = df['学分'].sum()
        total_weighted_points = df['权重分数'].sum()
        
        # 计算GPA
        gpa = total_weighted_points / total_credits if total_credits > 0 else 0
        
        # 保存到实例变量
        self.courses_data = df
        self.total_credits = total_credits
        self.weighted_points = total_weighted_points
        self.gpa = gpa
        
        return {
            'courses': df,
            'total_credits': total_credits,
            'total_weighted_points': total_weighted_points,
            'gpa': gpa,
            'course_count': len(df)
        }
    
    def display_results(self, results: Dict[str, Any]):
        """
        显示计算结果
        
        Args:
            results: 计算结果字典
        """
        print("\n" + "="*60)
        print("                    GPA计算结果")
        print("="*60)
        
        # 显示课程详情
        courses_df = results['courses']
        if '课程名称' in courses_df.columns:
            print(f"\n{'课程名称':<20} {'学分':<8} {'绩点':<8} {'权重分数':<10}")
            print("-" * 50)
            for _, row in courses_df.iterrows():
                print(f"{str(row['课程名称']):<20} {row['学分']:<8.1f} {row['绩点']:<8.2f} {row['权重分数']:<10.2f}")
        else:
            print(f"\n{'课程序号':<10} {'学分':<8} {'绩点':<8} {'权重分数':<10}")
            print("-" * 40)
            for i, row in courses_df.iterrows():
                print(f"课程{i+1:<7} {row['学分']:<8.1f} {row['绩点']:<8.2f} {row['权重分数']:<10.2f}")
        
        print("\n" + "-" * 60)
        print(f"课程总数: {results['course_count']} 门")
        print(f"总学分: {results['total_credits']:.1f}")
        print(f"总权重分数: {results['total_weighted_points']:.2f}")
        print(f"平均学分绩点(GPA): {results['gpa']:.4f}")
        print("="*60)
    
    def process_file(self, file_path: str) -> float:
        """
        处理Excel文件并计算GPA
        
        Args:
            file_path: Excel文件路径
            
        Returns:
            float: 计算得到的GPA
        """
        try:
            # 读取Excel文件
            df = self.read_excel_file(file_path)
            
            # 验证数据格式
            df_clean = self.validate_data_format(df)
            
            # 计算GPA
            results = self.calculate_gpa(df_clean)
            
            # 显示结果
            self.display_results(results)
            
            return results['gpa']
            
        except Exception as e:
            print(f"错误: {str(e)}")
            return 0.0

def main():
    """主函数"""
    parser = argparse.ArgumentParser(description="GPA计算器 - 从Excel文件计算学分绩点")
    parser.add_argument('file_path', help='Excel文件路径 (.xlsx格式)')
    parser.add_argument('--output', '-o', help='输出结果到文件（可选）')
    
    args = parser.parse_args()
    
    # 创建GPA计算器实例
    calculator = GPACalculator()
    
    # 处理文件
    gpa = calculator.process_file(args.file_path)
    
    # 如果指定了输出文件，保存结果
    if args.output and gpa > 0:
        try:
            with open(args.output, 'w', encoding='utf-8') as f:
                f.write(f"GPA计算结果\n")
                f.write(f"文件: {args.file_path}\n")
                f.write(f"总学分: {calculator.total_credits:.1f}\n")
                f.write(f"总权重分数: {calculator.weighted_points:.2f}\n")
                f.write(f"GPA: {calculator.gpa:.4f}\n")
            print(f"\n结果已保存到: {args.output}")
        except Exception as e:
            print(f"保存结果时出错: {str(e)}")

if __name__ == "__main__":
    main()
