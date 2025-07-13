# 动漫评分排名系统 (MzzbScore Edit)

一个专业的动漫评分数据处理和排名计算系统，支持从多个平台获取评分数据并生成排名报告。

## 📋 项目简介

本系统是一个面向对象的动漫评分数据处理工具，能够：

- 📊 从Excel文件读取动漫信息
- 🌐 整合多平台评分数据（Bangumi、Anilist、MyAnimelist、Filmarks）
- 🏆 计算综合评分和排名
- 🎨 保持Excel文件的样式和超链接

## ⭐ 主要功能

### 数据处理功能
- **智能数据过滤**: 自动过滤无效数据和特殊标记条目
- **综合评分计算**: 基于加权算法计算各平台的综合评分
- **排名生成**: 为每个平台和综合评分生成准确排名

### Excel处理功能
- **样式保持**: 保留原始Excel文件的格式和样式
- **超链接维护**: 自动保持和重新应用超链接
- **智能列管理**: 自动插入排名列并保持原有结构
- **数据完整性**: 确保数据处理过程中的完整性和准确性

## 🚀 快速开始

### 安装步骤

1. **克隆项目**
```bash
git clone <项目地址>
cd mzzbscore-edit
```

2. **安装依赖**
```bash
pip install -r requirements.txt
```

3. **准备数据文件**
   - 将动漫数据Excel文件放置在项目根目录，命名为 `mzzb.xlsx`
   - 确保Excel文件包含必要的列（原名、译名、各平台评分等）

4. **运行程序**
```bash
python main.py
```

## 📁 项目结构

```
mzzbscore-edit/
├── app/                    # 主应用程序模块
│   ├── config/            # 配置模块
│   │   ├── constants.py   # 常量定义
│   │   └── settings.py    # 配置管理
│   ├── core/              # 核心模块
│   │   ├── application.py # 主应用程序控制器
│   │   └── base.py        # 基础类定义
│   ├── models/            # 数据模型
│   │   └── data_models.py # 数据结构定义
│   ├── services/          # 服务层
│   │   ├── excel_service.py    # Excel处理服务
│   │   └── ranking_service.py  # 排名计算服务
│   └── utils/             # 工具模块
│       ├── exceptions.py  # 异常类定义
│       ├── logger.py      # 日志系统
│       └── validators.py  # 数据验证器
├── main.py                # 程序入口
├── requirements.txt       # 依赖项列表
├── mzzb.xlsx             # 数据文件（需要用户提供）
└── README.md             # 项目说明
```

## 💻 使用方法

### 方式一：使用编译好的EXE文件（推荐）

1. **下载程序**：从 [GitHub Releases](https://github.com/your-repo/releases) 页面下载最新版本的 `mzzbscore-edit.exe`
2. **准备数据**：下载配套的 `mzzb.xlsx` 模板文件，按格式填入动漫数据
3. **运行程序**：双击 `mzzbscore-edit.exe` 直接运行
4. **选择操作**：
   - 选择 "1" - 生成首月评分表格
   - 选择 "Q" - 退出程序

### 方式二：从源码运行

1. **启动程序**：运行 `python main.py`
2. **选择操作**：
   - 选择 "1" - 生成首月评分表格
   - 选择 "Q" - 退出程序

### 数据格式要求

Excel文件需要包含以下列：

#### 必需列
- `Notes`: 备注信息（用于过滤特殊条目）

#### 评分列（各平台）
- `Bangumi`: Bangumi评分
- `Anilist`: Anilist评分
- `MyAnimelist`: MyAnimeList评分
- `Filmarks`: Filmarks评分

#### 总评人数列
- `Bangumi_total`: Bangumi总评分人数
- `Anilist_total`: Anilist总评分人数
- `MyAnimelist_total`: MyAnimeList总评分人数
- `Filmarks_total`: Filmarks总评分人数

### 输出文件

程序会生成以下文件：
- `monthly_anime_scores.xlsx`: 包含排名信息的首月评分报告

## ⚙️ 配置选项

### 评分权重配置

系统使用加权平均计算综合评分，默认权重为：
- Bangumi: 50% (0.5)
- Anilist: 20% (0.2)
- MyAnimeList: 10% (0.1)
- Filmarks: 20% (0.2)

**计算公式**: `综合评分 = Bangumi×0.5 + Anilist×0.2 + MyAnimeList×0.1 + Filmarks×0.2`

### 数据过滤规则

系统会自动过滤以下条目：
- Notes列包含特殊标记的条目（如"*时长不足"、"*数据不足"等）
- 评分数据不完整的条目

### 样式配置

- **对齐方式**: 数据左对齐，垂直居中
- **列分组**: 按功能对列进行分组并使用不同背景色
- **边框**: 为所有数据单元格添加细边框