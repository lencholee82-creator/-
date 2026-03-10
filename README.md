# PDF 文件名整理工具

这是一个简单的 Python 工具，用于扫描指定文件夹中的 PDF 文件，并将文件名整理到 Excel 表格中。

## 功能

- 扫描指定目录下的 `.pdf` 文件
- 可选递归扫描子目录
- 生成 Excel 文件（包含文件名和完整路径）
- 仅依赖 Python 标准库

## 项目结构

```text
.
├── README.md
├── requirements.txt
└── src/
    └── pdf_inventory/
        ├── __init__.py
        ├── cli.py
        └── scanner.py
```

## 环境要求

- Python 3.10+

## 安装依赖

本项目仅使用标准库，无需安装第三方依赖：

```bash
pip install -r requirements.txt
```

## 运行方式

在仓库根目录执行：

```bash
PYTHONPATH=src python -m pdf_inventory.cli <要扫描的目录>
```

### 示例命令

1. 扫描当前目录下的 PDF 文件并输出到默认文件：

```bash
PYTHONPATH=src python -m pdf_inventory.cli ./documents
```

2. 指定输出 Excel 文件名：

```bash
PYTHONPATH=src python -m pdf_inventory.cli ./documents -o ./output/pdf_list.xlsx
```

3. 递归扫描子目录：

```bash
PYTHONPATH=src python -m pdf_inventory.cli ./documents -r
```

## 输出说明

生成的 Excel 默认文件名为 `pdf_filenames.xlsx`，表格列如下：

- `文件名`
- `完整路径`
