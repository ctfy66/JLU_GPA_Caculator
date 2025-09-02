# 🎓 GPA计算器

一个功能完整的学分绩点(GPA)计算工具，支持图形界面和命令行两种使用方式。可以从Excel文件读取课程信息并自动计算加权平均GPA。

![GitHub repo size](https://img.shields.io/github/repo-size/your-username/GPACaculator)
![Python version](https://img.shields.io/badge/python-3.6+-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

## ✨ 功能特性

### 🖥️ 图形用户界面 (推荐)
- 现代化的图形界面，用户友好
- 拖拽式文件选择
- 实时GPA计算和展示
- 详细成绩分析和统计
- 结果导出功能
- 完整的中文支持

### 💻 命令行界面
- 轻量级命令行工具
- 批处理友好
- 支持结果文件输出
- 适合自动化脚本

### 🔧 核心功能
- 📊 智能识别Excel列名（支持多种格式）
- 🧮 精确计算加权平均GPA：`GPA = sum(学分 × 绩点) / sum(学分)`
- 📈 提供详细的成绩分析报告
- ✅ 完善的数据验证和错误处理
- 🌍 完整的中文支持和本地化

## 🚀 快速开始

### 系统要求
- Python 3.6+
- Linux/Windows/macOS
- 对于GUI版本：需要图形界面支持

### 📦 安装依赖

1. **克隆项目**
```bash
git clone https://github.com/your-username/GPACaculator.git
cd GPACaculator
```

2. **安装Python依赖**
```bash
pip3 install -r requirements.txt
```

3. **Linux系统额外配置**（如果使用图形界面）
```bash
# Ubuntu/Debian系统
sudo apt update
sudo apt install python3-tk fonts-noto-cjk fonts-wqy-zenhei -y
fc-cache -fv

# CentOS/RHEL系统  
sudo yum install tkinter google-noto-cjk-fonts
```

4. **WSL2用户特别说明**
   - Windows 11用户：内置WSLg支持，无需额外配置
   - Windows 10用户：需要安装X服务器（如VcXsrv）

## 📖 使用方法

### 🖥️ 图形界面版本（推荐新手）

```bash
python3 gpa_gui.py
```

**使用步骤：**
1. 🔍 点击"浏览文件"选择Excel成绩文件
2. ✅ 程序自动验证文件格式
3. 🧮 点击"计算GPA"获得结果
4. 📊 查看详细的成绩分析
5. 💾 可选择保存结果到文本文件

### 💻 命令行版本（适合高级用户）

**基本使用：**
```bash
python3 gpa_calculator.py your_grades.xlsx
```

**保存结果到文件：**
```bash
python3 gpa_calculator.py your_grades.xlsx -o result.txt
```

**查看帮助信息：**
```bash
python3 gpa_calculator.py --help
```

## 📋 Excel文件格式要求

### 必需列
Excel文件必须包含以下信息（列名可以灵活变化）：

| 课程名称（可选） | 学分 | 绩点 |
|----------------|------|------|
| 高等数学A       | 4    | 3.7  |
| 线性代数        | 3    | 4.0  |
| 概率论与数理统计 | 3    | 3.3  |

### 🏷️ 支持的列名格式

**学分列：**
- `学分`、`学时`、`credit`、`credits`、`学分数`

**绩点列：**  
- `绩点`、`成绩`、`gpa`、`grade`、`绩点成绩`

**课程名称列（可选）：**
- `课程`、`课程名称`、`course`、`科目`、`课程名`

### 📝 数据要求
- 学分：必须 > 0 的数字
- 绩点：通常为 0-5 范围内的数字
- 程序会自动忽略包含空值的行
- 非数字数据会被过滤并提示

## 🎯 使用示例

### 示例文件
项目包含示例Excel文件，您可以直接测试：
```bash
python3 gpa_gui.py
# 然后选择 "全部成绩查询 (1).xlsx" 文件
```

### 计算结果示例
```
============================================================
                    GPA计算结果
============================================================

课程名称                 学分       绩点       权重分数      
--------------------------------------------------
高等数学A                4.0      3.70     14.80     
线性代数                 3.0      4.00     12.00     
概率论与数理统计             3.0      3.30     9.90      
C++程序设计              3.0      3.80     11.40     

------------------------------------------------------------
📚 课程总数: 4 门
📊 总学分: 13.0
📈 总权重分数: 48.10
🎯 平均学分绩点(GPA): 3.7000
============================================================

📋 成绩分析:
  优秀 (绩点≥4.0): 1 门
  良好 (3.0≤绩点<4.0): 3 门
  一般 (2.0≤绩点<3.0): 0 门
  待提升 (绩点<2.0): 0 门
```

## 🗂️ 项目文件结构

```
GPACaculator/
├── gpa_gui.py            # 🖥️ 图形界面版本（主推荐）
├── gpa_calculator.py     # 💻 命令行版本
├── requirements.txt      # 📦 Python依赖包
├── 全部成绩查询 (1).xlsx    # 📄 示例Excel文件
├── README.md            # 📖 项目说明文档
└── create_sample.py     # 🛠️ 示例文件生成脚本
```

## 🔧 系统环境配置

### Windows用户
```bash
# 安装依赖
pip install pandas openpyxl

# 直接运行
python gpa_gui.py
```

### Linux用户
```bash
# Ubuntu/Debian
sudo apt update
sudo apt install python3-tk fonts-noto-cjk fonts-wqy-zenhei
pip3 install pandas openpyxl

# CentOS/RHEL
sudo yum install tkinter google-noto-cjk-fonts
pip3 install pandas openpyxl
```

### macOS用户
```bash
# 使用Homebrew
brew install python-tk
pip3 install pandas openpyxl
```

## 🛠️ 故障排除

### 问题1：ModuleNotFoundError: No module named 'tkinter'
**解决方案：**
```bash
# Linux
sudo apt install python3-tk

# macOS  
brew install python-tk
```

### 问题2：中文字符显示为空白框
**解决方案：**
```bash
# 安装中文字体
sudo apt install fonts-noto-cjk fonts-wqy-zenhei fonts-wqy-microhei
fc-cache -fv
```

### 问题3：WSL2图形界面无法显示
**解决方案：**
- **Windows 11**：内置WSLg支持，重启WSL即可
- **Windows 10**：安装VcXsrv或X410，设置DISPLAY变量

### 问题4：Excel文件读取失败
**检查项目：**
- ✅ 文件是否为.xlsx格式
- ✅ 是否包含"学分"和"绩点"列
- ✅ 数据是否为数字格式
- ✅ 文件是否已关闭（Excel未占用）

## 🎨 界面预览

### 图形界面特性
- 🎯 清晰的标题和说明
- 📁 便捷的文件选择对话框
- 🧮 一键计算按钮
- 📊 实时结果显示
- 🎨 美观的成绩分析图表
- 💾 结果导出功能

### 成绩分析功能
- 📈 GPA等级评定（优秀/良好/及格/需努力）
- 📋 课程成绩分布统计
- 📊 学分权重分析
- 🏆 成绩趋势分析

## 📋 开发信息

### 技术栈
- **语言：** Python 3.6+
- **GUI框架：** tkinter
- **数据处理：** pandas
- **Excel处理：** openpyxl
- **字体支持：** Noto CJK、文泉驿字体

### 兼容性
- ✅ Windows 7/10/11
- ✅ Ubuntu 18.04+
- ✅ macOS 10.14+
- ✅ WSL2 (Windows 10/11)

## 🤝 贡献指南

欢迎提交Issue和Pull Request！

### 贡献方式
1. Fork 本项目
2. 创建功能分支 (`git checkout -b feature/amazing-feature`)
3. 提交更改 (`git commit -m '添加某个功能'`)
4. 推送到分支 (`git push origin feature/amazing-feature`)
5. 提交Pull Request

## 📜 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情。

## 📞 联系方式

如果您有任何问题或建议，请：
- 📧 提交Issue
- 💬 发起Discussion
- 🔗 Fork并改进项目

---
⭐ 如果这个项目对您有帮助，请给个星标支持一下！

## 更新日志

### v2.0.0
- ✨ 新增图形用户界面
- 🎨 完善的中文字体支持
- 📊 增强的成绩分析功能
- 🔧 改进的错误处理机制

### v1.0.0  
- 🚀 基础命令行GPA计算功能
- 📁 Excel文件读取支持
- 📈 基本的结果展示