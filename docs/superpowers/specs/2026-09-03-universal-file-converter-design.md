# 万能文件转换工具 — 设计规范

日期：2026-09-03
状态：已确认（自动模式决策）

## 1. 目标

开发一个面向日常办公场景的万能文件转换工具（Windows 桌面），覆盖办公常用文件类型的相互转换。提供图形界面（拖拽/选择文件、批量转换）与命令行两种使用方式。

## 2. 技术选型

- **语言/运行时**：Python 3.14（本机已装 `C:\Python314`）
- **GUI**：Tkinter（标准库，零额外依赖；ttk 主题）
- **CLI**：argparse（`python -m file_converter ...`）
- **核心依赖**：
  - Pillow — 图片转换（已装）
  - python-docx — docx 读写
  - openpyxl — xlsx 读写
  - python-pptx — pptx 读取
  - pypdf — PDF 合并/拆分/旋转/文本提取
  - PyMuPDF (fitz) — PDF→图片渲染
  - pywin32 — Office COM 自动化（docx/xlsx/pptx → PDF）
  - markdown — md→html
  - beautifulsoup4 — html→txt（已装）
  - PyYAML — yaml↔json（已装）
- **可选/回退**：无 MS Office 时，Office→PDF 回退 LibreOffice `soffice --headless --convert-to pdf`；两者都无则该功能在 GUI 中置灰并提示。

## 3. 架构

```
file_converter/
├── __init__.py          # 版本号
├── __main__.py          # CLI 入口 (argparse)
├── registry.py          # 转换器注册表：按 (源格式, 目标格式) 查找
├── converters/
│   ├── __init__.py      # 导入所有 converter 触发注册
│   ├── base.py          # Converter 基类 / @register 装饰器 / ConvertError
│   ├── images.py        # 图片互转、图片→PDF、PDF→图片
│   ├── documents.py     # docx/xlsx/pptx → pdf/txt/md/csv（Office COM + 库）
│   ├── pdf.py           # PDF 合并/拆分/旋转/提取文本
│   └── textdata.py      # csv/json/yaml/xml/md/html/txt + 编码转换
├── gui.py               # Tkinter 图形界面
└── deps.py              # 可选依赖探测，报告可用能力
tests/                   # pytest（只测纯库转换，不测 COM）
requirements.txt
README.md
```

### 注册表模式

每个转换函数用装饰器注册：

```python
@register(src=["png", "jpg", ...], dst=["pdf"])
def images_to_pdf(src: Path, dst: Path, **opts) -> None: ...
```

`registry.convert(src_path, dst_ext, **opts)` 查找并调用。GUI 据此动态生成"该文件可转成什么"的下拉选项；`registry.available_targets(src_ext)` 会过滤掉依赖缺失的 converter。

### 前端

- **GUI**：文件列表（添加文件/文件夹、清空）、目标格式下拉（随选中文件联动）、输出目录选择、进度条、结束汇总对话框。
- **CLI**：`python -m file_converter input.docx -t pdf [-o outdir]`，支持多文件与通配符；`--list` 打印支持的转换矩阵。

## 4. 转换矩阵

| 类别 | 源 | 目标 | 实现 |
|---|---|---|---|
| 文档 | docx | pdf | Word COM / LibreOffice 回退 |
| 文档 | docx | txt, md | python-docx |
| 文档 | xlsx | pdf | Excel COM / LibreOffice 回退 |
| 文档 | xlsx | csv | openpyxl（逐 sheet，多 sheet 时生成多个 csv） |
| 文档 | csv | xlsx | openpyxl |
| 文档 | pptx | pdf | PowerPoint COM / LibreOffice 回退 |
| 文档 | pptx | txt | python-pptx（提取各页文本） |
| 图片 | png/jpg/jpeg/webp/bmp/gif/tiff/ico | 互转 | Pillow |
| 图片 | 同上 | pdf | Pillow（多图合成一个 PDF） |
| PDF | pdf | png/jpg | PyMuPDF（逐页渲染，可选 DPI） |
| PDF | pdf | txt | pypdf 文本提取 |
| PDF | pdf | pdf（合并/拆分/旋转） | pypdf（特殊操作，CLI/GUI 单独入口） |
| 数据 | csv | json, yaml | csv + json/yaml |
| 数据 | json | csv, yaml | 记录数组→表 |
| 数据 | yaml | json, csv | PyYAML |
| 数据 | xml | json | xml.etree → dict → json |
| 文本 | md | html | markdown |
| 文本 | html | txt | beautifulsoup4 |
| 文本 | txt | 编码转换（utf-8/gbk/big5/shift_jis...） | 重编码写回 |

## 5. 数据流

GUI/CLI 收集 (文件列表, 目标格式, 选项) → 逐文件调用 `registry.convert` → 每个文件独立 try/except，失败记录原因继续 → 结束输出成功/失败汇总。输出文件默认写到所选输出目录，文件名 = 原 stem + 新扩展名；重名自动加 `(1)` 后缀，绝不覆盖。

## 6. 错误处理

- 自定义 `ConvertError`，converter 内捕获底层异常并包装为带中文说明的 `ConvertError`。
- 依赖缺失：该 converter 注册时标记 `requires`，`deps.py` 启动探测，不可用的转换不出现在 GUI 选项中，CLI 调用时报清晰错误。
- COM 转换失败（Office 未安装/被占用）自动回退 LibreOffice，再失败则报错。

## 7. 测试

pytest 覆盖纯库 converter：
- 图片互转、图片→PDF
- csv↔xlsx、csv↔json↔yaml、xml→json
- md→html、html→txt、编码转换（GBK↔UTF-8）
- PDF 合并/拆分/旋转/提取文本（用 pypdf 现场生成测试 PDF）
- docx→txt、pptx→txt

不测：Office COM / LibreOffice 路径（环境依赖）、GUI。

## 8. 验收标准

1. `pip install -r requirements.txt` 后 `pytest` 全绿。
2. CLI：`python -m file_converter --list` 输出转换矩阵；对 docx/xlsx/csv/png/pdf/md 各执行一次真实转换成功。
3. GUI 可启动，添加文件后目标格式下拉正确联动，批量转换有汇总。
