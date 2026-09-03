# 万能文件转换工具 (File Converter)

一个 Windows 上的本地文件转换工具，提供 **图形界面（Tkinter）** 和 **命令行（CLI）** 两种用法，覆盖 Office 文档、图片、PDF、文本与数据格式之间的常见互转，以及 PDF 合并 / 拆分 / 旋转。

- 纯本地转换，不上传任何文件
- 批量转换：一次传入多个文件，逐个处理，失败不影响其他文件
- 输出自动防覆盖：目标文件已存在时自动改名（追加 ` (1)`、` (2)`…… 后缀）
- 文本编码自动探测：依次尝试 `utf-8-sig`、`gbk`、`big5`、`shift_jis`、`latin-1`

## 环境要求

- Python 3.14（Windows）
- Office→PDF 转换需要本机安装 **MS Office**（Word/Excel/PowerPoint，走 COM 接口）或 **LibreOffice**

## 安装

```bash
pip install -r requirements.txt
```

## 图形界面（GUI）

```bash
python -m file_converter --gui
```

界面支持：添加文件 / 文件夹、按选中文件自动过滤可用目标格式、选择输出目录、调整 DPI（PDF 转图片）与文本编码，以及合并 / 拆分 / 旋转 PDF。

## 命令行（CLI）

`convert` 是默认子命令，可以省略。

```bash
# 列出全部支持的转换（[*] 表示依赖缺失暂不可用）
python -m file_converter list

# 格式转换（可一次传多个文件；-o 指定输出目录，默认同源文件目录）
python -m file_converter convert a.docx b.docx -t pdf -o out
python -m file_converter a.docx -t txt            # 省略 convert 也可以

# 合并多个 PDF（至少 2 个）
python -m file_converter merge a.pdf b.pdf c.pdf -o merged.pdf

# 按页码拆分 PDF（输出多份：xxx_p1-3.pdf、xxx_p5.pdf）
python -m file_converter split in.pdf --pages 1-3,5 -o outdir

# 旋转 PDF（角度仅支持 90 / 180 / 270）
python -m file_converter rotate in.pdf --angle 90 -o rotated.pdf
```

### convert 可选参数

| 参数 | 说明 |
| --- | --- |
| `-t, --to` | 目标格式（必填），如 `pdf` / `jpg` / `xlsx` |
| `-o, --outdir` | 输出目录，默认与源文件同目录 |
| `--src-encoding` | 源文本编码，默认自动探测 |
| `--dst-encoding` | txt 输出编码，默认 `utf-8` |
| `--dpi` | PDF 转图片的 DPI，默认 `150` |

## 支持的转换

| 源格式 | 目标格式 | 说明 |
| --- | --- | --- |
| docx | txt, md | Word 提取纯文本 / Markdown（保留标题层级） |
| xlsx | csv | Excel 转 CSV（多工作表时每个工作表输出一份 `文件名_表名.csv`） |
| csv | xlsx | CSV 转 Excel |
| pptx | txt | PPT 提取纯文本（按页分节） |
| docx, xlsx, pptx | pdf | Office→PDF（见下方引擎说明） |
| png, jpg, jpeg, webp, bmp, gif, tiff, ico | 同上互转 | 图片格式互转（Pillow） |
| png, jpg, jpeg, webp, bmp, gif, tiff, ico | pdf | 图片转 PDF |
| pdf | txt | PDF 提取纯文本 |
| pdf | png, jpg, jpeg | PDF 逐页转图片（多页时输出 `文件名_p1.png`、`文件名_p2.png`……） |
| csv | json, yaml, yml | 表格数据转 JSON / YAML（首行为表头） |
| json | csv, yaml, yml | JSON 转 CSV（要求对象数组）/ YAML |
| yaml, yml | json, csv | YAML 转 JSON / CSV |
| xml | json | XML 转 JSON（属性带 `@` 前缀，文本节点为 `#text`） |
| md, markdown | html | Markdown 转 HTML（支持表格、代码块） |
| html, htm | txt | HTML 提取纯文本 |
| txt | txt | 文本编码转换（配合 `--src-encoding` / `--dst-encoding`） |

PDF 工具（`merge` / `split` / `rotate`）见上方 CLI 用法。

## Office→PDF 引擎说明

`docx` / `xlsx` / `pptx` → `pdf` 按以下顺序选择引擎：

1. **MS Office（COM）**：本机装有 Microsoft Office 时优先使用，通过 Word / Excel / PowerPoint 的 COM 接口导出，保真度最高（首次调用需启动 Office，可能耗时数秒）。
2. **LibreOffice（回退）**：未检测到 MS Office 时，自动查找 `soffice`（PATH 或 `Program Files\LibreOffice`）以无头模式转换。

两者都不可用时该转换会报 `[失败]`，并提示安装 MS Office 或 LibreOffice。`python -m file_converter list` 会显示当前实际使用的引擎。

## 常见问题

**Q：命令行输出的中文在终端里显示为乱码？**
A：Windows 控制台默认使用 GBK 代码页，而程序按 UTF-8 输出时可能出现显示乱码。这只是**显示问题**，写入文件的内容不受影响。可在运行前执行 `chcp 65001` 切换控制台到 UTF-8，或设置环境变量 `PYTHONIOENCODING=utf-8`。

**Q：转换出的 CSV 用 Excel 打开中文乱码？**
A：程序输出的 CSV 统一使用 **UTF-8 带 BOM（utf-8-sig）** 编码，双击用 Excel 打开即可正常显示中文。若用其他工具读取，注意跳过前 3 字节的 BOM。

**Q：读取旧的 GBK 编码文本文件失败？**
A：文本类输入默认按 `utf-8-sig → gbk → big5 → shift_jis → latin-1` 顺序自动探测编码，一般无需干预；探测失败时可用 `--src-encoding gbk` 显式指定。

**Q：输出目录里已有同名文件会被覆盖吗？**
A：不会。所有输出路径都经过防覆盖处理，已存在时自动追加 ` (1)`、` (2)`…… 后缀；PDF 拆分、多页 PDF 转图片等多产物场景同样适用。

**Q：某条转换标记为 `[*]` 不可用？**
A：表示对应第三方库未安装（如 PyYAML、PyMuPDF）。重新执行 `pip install -r requirements.txt` 即可；Office→PDF 额外需要 MS Office 或 LibreOffice。
