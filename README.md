# PMC Pro轻松综述 从PMC生成一个ENDNOTE文献格式的word

PMC Pro是一个一站式解决方案，用于下载PMC文章并生成带有EndNote兼容引用的Word文档。该工具接受PMID作为输入，自动下载PMC全文，提取参考文献，并生成格式化的Word文档和RIS文件。

## 功能特点

- **一键下载PMC文章**：只需提供PMID，自动下载PMC全文XML
- **参考文献提取**：从PMC文章中提取所有参考文献，包括作者、标题、期刊等信息
- **生成RIS文件**：创建EndNote兼容的RIS格式引用文件
- **创建Word文档**：生成带有精确引用标记的Word文档
- **唯一引用标识符**：使用PMCID创建全局唯一的引用ID，避免合并多篇文章时的引用冲突
- **主RIS文件管理**：自动将所有引用合并到一个主RIS文件，方便一次性导入EndNote
- **重复引用检测**：自动检测并清理重复的引用记录，保留最新版本

## 安装说明

### 前提条件

- Python 3.6或更高版本
- pip包管理器

### 安装步骤

1. 克隆仓库：

```bash
git clone https://github.com/yourusername/pmc-pro.git
cd pmc-pro
```

2. 安装依赖项：

```bash
pip install -r requirements.txt
```

或手动安装所需包：

```bash
pip install loguru lxml python-docx beautifulsoup4 requests tqdm
```

## 使用方法

### 基本用法

```bash
# 处理单个PMID
python pmc_pro.py --pmid 32066954

# 处理多个PMID
python pmc_pro.py --pmids 32066954,31978945,32376714

# 从文件读取PMID列表
python pmc_pro.py --file pmids.txt
```

### 高级选项

```bash
# 指定输出目录
python pmc_pro.py --pmid 32066954 --output my_articles

# 保留中间XML文件
python pmc_pro.py --pmid 32066954 --keep-xml

# 显示详细输出
python pmc_pro.py --pmid 32066954 --verbose

# 清理主RIS文件中的重复记录
python pmc_pro.py --clean-ris

# 清理指定RIS文件
python pmc_pro.py --clean-ris --ris-file path/to/your/file.ris
```

### EndNote集成流程

1. 运行脚本处理一个或多个PMID
2. 在EndNote中导入生成的主RIS文件：
   - 打开EndNote
   - 选择 File > Import > File...
   - 选择 `output/endnote_import/master_references.ris` 文件
   - Import Option 选择 'Reference Manager (RIS)'
   - 点击 Import 按钮
3. 确认Label字段包含带有PMCID的唯一标识符
4. 在EndNote偏好设置中启用Label匹配：
   - 选择 Edit > Preferences > Formatting > Temporary Citations
   - 勾选 "Use field instead of record number"
   - 选择 "Label" 作为字段
5. 打开生成的Word文档
6. 使用EndNote的'Update Citations and Bibliography'功能

## 文件结构

- `pmc_pro.py` - 主脚本，处理命令行参数并协调整个工作流程
- `PMC_converter.py` - 负责下载PMC文章
- `generate_ris.py` - 从PMC XML提取参考文献并生成RIS文件
- `generate_word.py` - 创建带有引用标记的Word文档
- `requirements.txt` - 项目依赖项列表

## 输出文件

脚本会在指定的输出目录（默认为`pmc_output`）中创建以下文件和目录：

- `PMC{ID}/` - 每篇文章的专用目录，以PMCID命名
  - `PMC{ID}.xml` - PMC文章的XML文件
  - `PMC{ID}.ris` - 该文章的参考文献RIS文件
  - `PMC{ID}.docx` - 带有引用标记的Word文档
- `endnote_import/` - EndNote导入相关文件
  - `master_references.ris` - 包含所有文章参考文献的主RIS文件
  - `README.txt` - EndNote导入说明
- `processing_results.json` - 处理结果摘要
- `pmc_pro.log` - 详细日志文件

## 技术细节

### 唯一引用标识符

为了解决多篇文章合并时引用ID冲突的问题，PMC Pro使用以下格式创建全局唯一的引用标识符：

- RIS文件中: `LB  - ^PMC7096285_R1$`
- Word文档中: `{^PMC7096285_R1$}`

这样，即使不同文章使用相同的引用ID（如R1、R2），通过添加PMCID前缀，每个引用都有一个唯一的标识符。

### 主RIS文件管理

PMC Pro会自动将所有处理过的文章的参考文献合并到一个主RIS文件中，并提供重复检测和清理功能：

1. 每次处理新文章时，将其参考文献追加到主RIS文件
2. 自动检测重复的引用记录（基于LB标识符）
3. 保留每个唯一标识符的最新记录
4. 提供手动清理选项

## 依赖项

- `loguru` - 高级日志记录
- `lxml` - XML处理
- `python-docx` - Word文档生成
- `beautifulsoup4` - HTML解析
- `requests` - HTTP请求
- `tqdm` - 进度条显示

