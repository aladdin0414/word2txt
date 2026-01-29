# word2txt

一个简单的本地脚本集合，用于：

- 将 `.docx` 转为 `.txt`
- 将 `.pdf` 转为 `.txt`（依赖可选安装，自动选择解析引擎）
- 将一个目录下的多个 `.txt` 合并为一个文件

## 目录结构约定

默认按项目目录下的以下结构工作（也支持命令行参数指定输入/输出目录）：

```text
word2txt/
  word/        # 放 .docx
  pdf/         # 放 .pdf
  txt/         # 输出 .txt 或准备合并的 .txt
  merge/       # 合并输出目录
  word2txt.py
  pdf2txt.py
  txt-merge.py
```

## 环境要求

- Python 3.10+（建议 3.11/3.12）

### PDF 转换依赖（四选一）

`pdf2txt.py` 会自动检测并优先选择可用引擎：

1. PyMuPDF（推荐，效果通常更好）
2. pdfminer.six
3. pypdf
4. PyPDF2

安装任意一个即可，例如：

```bash
pip install pymupdf
```

或：

```bash
pip install pdfminer.six
```

## 使用方法

### 1) DOCX 转 TXT

将 `.docx` 放到 `word/` 目录，然后运行：

```bash
python3 word2txt.py
```

自定义输入/输出目录：

```bash
python3 word2txt.py --input /path/to/docx_dir --output /path/to/txt_dir
```

输出文件名规则：`xxx.docx` → `xxx.txt`

### 2) PDF 转 TXT

将 `.pdf` 放到 `pdf/` 目录，然后运行：

```bash
python3 pdf2txt.py
```

自定义输入/输出目录：

```bash
python3 pdf2txt.py --input /path/to/pdf_dir --output /path/to/txt_dir
```

输出文件名规则：`xxx.pdf` → `xxx.txt`

说明：

- 脚本会打印当前使用的解析引擎（如 `pymupdf`、`pdfminer` 等）
- 单个 PDF 解析失败会被跳过，继续处理其它文件，并统计“已转换/跳过”数量

### 3) 合并 TXT

将需要合并的 `.txt` 放到 `txt/` 目录，然后运行：

```bash
python3 txt-merge.py
```

默认输出到 `merge/merged.txt`，合并顺序按文件名排序。

编码说明：

- 会依次尝试 `utf-8`、`utf-8-sig`、`gb18030`
- 若仍无法解码，会用 `utf-8` 以替换错误字符的方式读取
