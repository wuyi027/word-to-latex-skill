# Word to LaTeX Skill / Word 转 LaTeX 技能

Convert Word `.docx` manuscripts into compile-ready LaTeX projects for Codex, Claude Code, CI scripts, and direct command-line use.

将 Word `.docx` 论文/报告/学位论文转换为可编译的 LaTeX 工程，适用于 Codex、Claude Code、CI 脚本或命令行直接调用。

## Features / 功能

- Preserve document order for headings, paragraphs, tables, inline images, and many embedded graphics.
- Convert common Word OMML equations to LaTeX math, including fractions, scripts, radicals, delimiters, large operators, accents, and simple matrices.
- Rebuild Word tables as editable `booktabs` / `longtable` LaTeX tables.
- Extract images into `figures/` and insert referenced graphics into `main.tex`.
- Rasterize referenced EMF/WMF graphics into uniquely named PNG files with ImageMagick `magick` when available.
- Perform light Elsevier `elsarticle` cleanup: move leading `Abstract:` and `Keywords:` into `frontmatter`, remove duplicate body title, and normalize numbered headings.
- Detect reference sections and write both `references.tex` and a conservative `references.bib`.
- Optionally preprocess MathType equations through Word/MathType `Alt + \` before conversion.
- Optionally compile with `xelatex`, `pdflatex`, or `lualatex`, and write `compile.log`.
- Support JSON output for automation.

---

- 保留标题、段落、表格、行内图片和多数嵌入图形的文档顺序。
- 将常见 Word OMML 公式转换为 LaTeX 数学表达式，包括分式、上下标、根号、括号、大算子、重音符号和简单矩阵。
- 将 Word 表格重建为可编辑的 `booktabs` / `longtable` LaTeX 表格。
- 将图片提取到 `figures/`，并在 `main.tex` 中插入实际引用的图像。
- 如果系统安装了 ImageMagick `magick`，会自动把被引用的 EMF/WMF 图形转为唯一命名的 PNG，避免 XeLaTeX BoundingBox 错误和文件名覆盖。
- 对 Elsevier `elsarticle` 模板做轻量整理：把开头 `Abstract:` / `Keywords:` 放进 `frontmatter`，移除正文重复标题，并规范编号标题。
- 检测参考文献区域，生成 `references.tex` 和保守的 `references.bib`。
- 可选：在转换前通过 Word/MathType 的 `Alt + \` 批量将 MathType 公式转成 LaTeX 文本。
- 可选：使用 `xelatex`、`pdflatex` 或 `lualatex` 编译，并写入 `compile.log`。
- 支持 JSON 输出，方便自动化流程。

## Directory Layout / 目录结构

```text
word-to-latex-skill/
├── SKILL.md
├── README.md
├── requirements.txt
├── scripts/
│   └── word_to_latex.py
├── assets/
│   └── default_template.tex
└── references/
    └── README.md
```

## Requirements / 环境依赖

Required for conversion:

转换必需：

```bash
python -m pip install -r requirements.txt
```

Python packages:

Python 包：

- `lxml`
- `python-docx`
- `pylatexenc`

Required for PDF compilation:

编译 PDF 需要安装 LaTeX 发行版：

- TeX Live, MacTeX, or MiKTeX
- Recommended: `latexmk` + `xelatex`
- Chinese or mixed-language documents should use `xelatex`

Optional tools:

可选工具：

- ImageMagick `magick`, used to rasterize referenced EMF/WMF graphics.
- Microsoft Word + MathType, only for `--preprocess-mathtype`.
- `pandoc`, currently only checked as an optional fallback tool.

## Dependency Check / 依赖检查

```bash
python scripts/word_to_latex.py --check-deps
```

JSON output:

JSON 输出：

```bash
python scripts/word_to_latex.py --check-deps --json
```

Install missing Python packages with an explicit prompt:

显式确认后安装缺失的 Python 包：

```bash
python scripts/word_to_latex.py --check-deps --install-missing
```

Non-interactive mode:

非交互模式：

```bash
python scripts/word_to_latex.py --check-deps --install-missing --yes
```

## Usage / 使用方式

### Basic conversion / 基础转换

```bash
python scripts/word_to_latex.py manuscript.docx \
  -o build/latex_project \
  --json
```

### Use a custom template / 使用自定义模板

```bash
python scripts/word_to_latex.py manuscript.docx \
  -t template.tex \
  -o build/latex_project \
  --json
```

The template should contain:

模板建议包含以下占位符：

```tex
%%TITLE%%
%%CONTENT%%
%%BIBLIOGRAPHY%%
```

### Convert and compile / 转换并编译

```bash
python scripts/word_to_latex.py manuscript.docx \
  -t template.tex \
  -o build/latex_project \
  --engine xelatex \
  --compile \
  --json
```

### Elsevier template / Elsevier 模板

Use an `elsarticle` template with the same placeholders:

使用带占位符的 `elsarticle` 模板：

```tex
\documentclass[preprint,12pt]{elsarticle}
\usepackage{graphicx}
\usepackage{booktabs}
\usepackage{longtable}
\usepackage{amsmath,amssymb}
\usepackage{hyperref}
\usepackage{xeCJK}
\graphicspath{{figures/}}
\journal{Elsevier}
\begin{document}
\begin{frontmatter}
\title{%%TITLE%%}
\author{}
\address{}
\end{frontmatter}
%%CONTENT%%
%%BIBLIOGRAPHY%%
\end{document}
```

When an Elsevier template is detected, the converter tries to move leading `Abstract:` and `Keywords:` paragraphs into `frontmatter`.

检测到 Elsevier 模板时，转换器会尝试将正文开头的 `Abstract:` 和 `Keywords:` 移入 `frontmatter`。

### MathType preprocessing / MathType 公式预处理

If formulas were created with MathType and saved as images or OLE objects, the `.docx` XML usually does not contain LaTeX. In that case, run a non-destructive Word/MathType preprocessing pass first:

如果公式由 MathType 创建并以图片或 OLE 对象保存在 Word 中，`.docx` XML 通常不包含可直接读取的 LaTeX。此时可以先执行非破坏性的 Word/MathType 预处理：

```bash
python scripts/word_to_latex.py manuscript.docx \
  -t elsevier_template.tex \
  -o build/latex_project \
  --preprocess-mathtype \
  --engine xelatex \
  --compile \
  --json
```

This creates `*_mathtype_preprocessed.docx` in the output directory, sends MathType's `Alt + \` shortcut to selected embedded formula objects, then converts the copied file.

该步骤会在输出目录中创建 `*_mathtype_preprocessed.docx`，对选中的嵌入公式对象发送 MathType 的 `Alt + \` 快捷键，然后再转换这个中间文件。

Useful options:

常用选项：

```bash
--mathtype-visible      # show Word during preprocessing / 显示 Word 窗口
--mathtype-wait 1.5     # wait longer after each Alt+\ / 每个公式后等待更久
```

This feature requires Microsoft Word, MathType, and an interactive Windows session.

该功能依赖 Microsoft Word、MathType 和可交互的 Windows 桌面会话。

## Install as a Codex Skill / 安装为 Codex Skill

Repository-level install:

仓库级安装：

```bash
mkdir -p .agents/skills
cp -R word-to-latex-skill .agents/skills/word-to-latex
python -m pip install -r .agents/skills/word-to-latex/requirements.txt
python .agents/skills/word-to-latex/scripts/word_to_latex.py --check-deps
```

User-level install:

用户级安装：

```bash
mkdir -p "$HOME/.agents/skills"
cp -R word-to-latex-skill "$HOME/.agents/skills/word-to-latex"
python -m pip install -r "$HOME/.agents/skills/word-to-latex/requirements.txt"
```

Then ask Codex:

然后可以让 Codex 执行：

```text
Use the word-to-latex skill to convert manuscript.docx into LaTeX and compile with xelatex.
```

## Install in Claude Code / 安装到 Claude Code

Project-level install:

项目级安装：

```bash
mkdir -p .claude/skills
cp -R word-to-latex-skill .claude/skills/word-to-latex
python -m pip install -r .claude/skills/word-to-latex/requirements.txt
```

Personal install:

个人级安装：

```bash
mkdir -p "$HOME/.claude/skills"
cp -R word-to-latex-skill "$HOME/.claude/skills/word-to-latex"
python -m pip install -r "$HOME/.claude/skills/word-to-latex/requirements.txt"
```

## Output / 输出文件

Typical output:

典型输出：

```text
build/latex_project/
├── main.tex
├── main.pdf              # when compilation succeeds / 编译成功时生成
├── compile.log           # when --compile is used / 使用 --compile 时生成
├── references.tex
├── references.bib
├── manuscript_mathtype_preprocessed.docx  # only with --preprocess-mathtype
└── figures/
    ├── figure_001.png
    └── figure_002_raster.png
```

JSON summary example:

JSON 摘要示例：

```json
{
  "title": "Example Manuscript",
  "tex": "build/latex_project/main.tex",
  "figures": 3,
  "rasterized_vector_graphics": 1,
  "tables": 2,
  "references": 18,
  "compile": {
    "compiled": true,
    "pdf": "build/latex_project/main.pdf",
    "log": "build/latex_project/compile.log"
  }
}
```

## Troubleshooting / 故障排查

### Missing dependency: `lxml` / 缺少 `lxml`

```bash
python scripts/word_to_latex.py --check-deps --install-missing
```

or:

或：

```bash
python -m pip install -r requirements.txt
```

### No LaTeX compiler found / 未找到 LaTeX 编译器

Install TeX Live, MacTeX, or MiKTeX, then verify:

安装 TeX Live、MacTeX 或 MiKTeX，然后验证：

```bash
xelatex --version
```

### EMF/WMF graphics fail / EMF/WMF 图像编译失败

Install ImageMagick and make `magick` available in `PATH`. The converter rasterizes referenced EMF/WMF files automatically when possible.

安装 ImageMagick，并确保 `magick` 在 `PATH` 中。转换器会尽可能自动转换被引用的 EMF/WMF 图形。

### MathType preprocessing does nothing / MathType 预处理没有效果

Use:

使用：

```bash
--mathtype-visible --mathtype-wait 1.5
```

Confirm that Word and MathType are installed, and that `Alt + \` works manually for the selected formulas.

确认 Word 和 MathType 已安装，并且手动选中公式后 `Alt + \` 可以正常转换。

## Limitations / 已知限制

- `.doc` files must be converted to `.docx` first.
- OMML equation conversion covers common scientific formulas, but very complex equation layouts may still need manual cleanup.
- MathType equations stored only as images/OLE objects do not expose LaTeX in `.docx`; use `--preprocess-mathtype` or convert them manually in Word/MathType first.
- Scanned or image-only equations require OCR or manual rewriting.
- Complex merged cells are flattened and may require manual cleanup.
- EndNote, Zotero, or Mendeley field-code metadata is preserved conservatively, not deeply parsed.
- Generated BibTeX entries use `@misc` unless richer parsing is added later.

---

- `.doc` 文件需要先转换为 `.docx`。
- OMML 公式转换覆盖常见科学公式，但复杂公式布局可能仍需人工整理。
- 如果 MathType 公式只以图片/OLE 对象存在，`.docx` 中不会暴露 LaTeX；请使用 `--preprocess-mathtype`，或先在 Word/MathType 中手动转换。
- 扫描版或纯图片公式需要 OCR 或手工重写。
- 复杂合并单元格会被展平，可能需要后处理。
- EndNote、Zotero、Mendeley 字段代码会保守保留文本，不做深度语义解析。
- 生成的 BibTeX 默认使用 `@misc`，后续可扩展更精细的解析。
