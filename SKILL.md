---
name: word-to-latex
version: 1.0.0
description: Convert Word .docx manuscripts into compile-ready LaTeX projects. Use when the user asks to transform Word to LaTeX, extract figures, redraw tables, extract references, apply a LaTeX template, or compile/validate the output PDF.
---

# Word to LaTeX Skill

Use this skill to convert `.docx` documents into a LaTeX project that can be compiled and inspected. The workflow is designed for manuscripts, reports, theses, and papers where the user needs figures extracted, tables rebuilt as LaTeX, references separated, and the final `.tex` validated by compilation.

## When to use

Use this skill when the task includes any of the following:

- Convert Word, DOCX, manuscript, report, paper, or thesis to LaTeX.
- Extract images from figures in a Word document.
- Rebuild Word tables as LaTeX `tabular` or `longtable` instead of screenshots.
- Extract references or bibliography from Word and create LaTeX bibliography content.
- Apply an existing LaTeX template and compile the result.
- Validate that the generated LaTeX compiles without errors.

Do not use this skill for scanned PDFs or image-only documents unless OCR or manual transcription is separately requested.

## Files

- `scripts/word_to_latex.py`: command-line converter used by Codex or a shell.
- `assets/default_template.tex`: fallback template with `%%TITLE%%`, `%%CONTENT%%`, and `%%BIBLIOGRAPHY%%` placeholders.
- `requirements.txt`: Python dependencies.
- `references/README.md`: implementation notes and template contract.

## Required environment

Install Python dependencies before running:

```bash
python -m pip install -r requirements.txt
```

For compilation, install a LaTeX distribution with at least one of:

- `latexmk` + `xelatex` recommended;
- `xelatex` for Chinese/mixed-language documents;
- `pdflatex` for pure English documents;
- `lualatex` if the template requires it.


## Dependency check before conversion

Before conversion or compilation, especially in a fresh environment, run:

```bash
python scripts/word_to_latex.py --check-deps
```

If Python packages are missing, ask the user for permission before installing them, or use the built-in permission prompt:

```bash
python scripts/word_to_latex.py --check-deps --install-missing
```

For non-interactive automation where the user has already granted permission, use:

```bash
python scripts/word_to_latex.py --check-deps --install-missing --yes
```

Do not silently install OS-level LaTeX distributions. If `xelatex`, `pdflatex`, or `lualatex` is missing, tell the user to install TeX Live, MacTeX, or MiKTeX for their operating system.

## Optional MathType preprocessing

If Word formulas were created with MathType and saved as OLE/drawing objects, first convert them inside Word/MathType before the normal DOCX-to-LaTeX conversion:

```bash
python scripts/word_to_latex.py manuscript.docx \
  -t elsevier_template.tex \
  -o build/latex_project \
  --preprocess-mathtype \
  --engine xelatex \
  --compile \
  --json
```

This creates a non-destructive intermediate copy named `*_mathtype_preprocessed.docx` in the output directory. It uses Microsoft Word COM automation and sends MathType's `Alt + \` shortcut to selected embedded formula objects. It requires Microsoft Word, MathType, and an interactive Windows session. If the shortcut does not work in the user's MathType setup, run Word visibly with `--mathtype-visible` and increase `--mathtype-wait`.

## Standard Codex workflow

1. Locate the input `.docx` and optional `.tex` template.
2. Create an isolated output directory, for example `build/latex_project`.
3. Run the converter script.
4. Inspect `main.tex`, `figures/`, `references.tex`, `references.bib`, and `compile.log` if compilation was requested.
5. If compilation fails, repair the generated LaTeX and rerun compilation until it passes, unless the environment lacks a LaTeX compiler.
6. Report the output path, number of extracted figures, tables, references, and compile status.

## Command examples

Convert with the default template:

```bash
python scripts/word_to_latex.py manuscript.docx -o build/latex_project --json
```

Convert with a user-provided LaTeX template:

```bash
python scripts/word_to_latex.py manuscript.docx \
  -t template.tex \
  -o build/latex_project \
  --json
```

Convert and compile with XeLaTeX:

```bash
python scripts/word_to_latex.py manuscript.docx \
  -t template.tex \
  -o build/latex_project \
  --engine xelatex \
  --compile \
  --json
```

## Template contract

The template should contain these placeholders:

```tex
%%TITLE%%
%%CONTENT%%
%%BIBLIOGRAPHY%%
```

If a user supplies a journal template without placeholders, insert `%%CONTENT%%` after `\begin{document}` and before the bibliography section. Put `%%BIBLIOGRAPHY%%` where references should appear. Keep required class/package declarations from the template.

## Conversion behavior

- Headings are mapped to `\section`, `\subsection`, `\subsubsection`, or `\paragraph`.
- Embedded Word media are extracted from `word/media/` into `figures/`.
- Inline images become LaTeX `figure` environments with `\includegraphics`.
- Captions beginning with `Figure`, `Fig.`, or `图` are attached to the nearest preceding image when possible.
- Word tables are rebuilt as `booktabs` tables; long tables are emitted as `longtable`.
- Common Word OMML equations are converted to LaTeX math, including inline/display formulas, fractions, scripts, radicals, large operators, delimiters, accents, and simple matrices.
- If an Elsevier `elsarticle` template is used, duplicate body titles plus leading `Abstract:` and `Keywords:` paragraphs are cleaned into `frontmatter` when possible.
- Referenced EMF/WMF graphics are rasterized to uniquely named PNG files with ImageMagick `magick` when available, avoiding XeLaTeX BoundingBox errors and filename collisions.
- Captions beginning with `Table` or `表` are attached to the nearest table when possible.
- A detected `References`, `Bibliography`, `参考文献`, `参考资料`, or `文献` heading starts the reference extraction region.
- References are written to both `references.tex` and `references.bib`. The generated `.bib` is conservative and uses `@misc` entries unless richer parsing is added.

## Quality checks Codex should perform after conversion

- Confirm figures exist under `figures/` and are referenced from `main.tex`.
- Confirm Word tables are represented as editable LaTeX tables, not screenshots.
- Confirm the references section is not duplicated in the body.
- Confirm labels are unique enough for figures and tables.
- Compile with the requested engine. Prefer `latexmk` when available.
- If LaTeX compilation fails, read `compile.log`, fix the first real LaTeX error, and recompile.
- For Chinese documents, prefer `xelatex` and keep `xeCJK` in the template.

## Known limitations

- `.doc` must be converted to `.docx` before using this skill.
- OMML equation conversion covers common scientific formulas, but very complex Word equation layouts may still need manual cleanup after conversion.
- MathType equations saved in Word as images or OLE/drawing objects do not expose LaTeX in the `.docx` XML. If MathType supports it, use Word/MathType's `Alt + \` conversion on those formulas before running this skill so the document contains LaTeX text or editable equation objects.
- Complex merged cells are flattened; manual cleanup may be needed for highly formatted tables.
- EndNote/Mendeley field-code parsing is not deeply semantic; final bibliography style should be verified.
