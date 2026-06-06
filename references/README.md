# Word to LaTeX Skill Notes

This skill converts `.docx` files into a LaTeX project. It extracts embedded media from the DOCX zip, converts tables to `booktabs`/`longtable`, detects references sections, creates both `references.tex` and `references.bib`, and optionally compiles the project.

Recommended LaTeX engine:
- `xelatex` for Chinese or mixed-language documents.
- `pdflatex` for pure English documents.
- `latexmk` if available, because it handles multiple passes.

Template contract:
- `%%TITLE%%` will be replaced by the inferred title or file stem.
- `%%CONTENT%%` will be replaced by converted body content.
- `%%BIBLIOGRAPHY%%` will be replaced by either `\input{references.tex}` or a BibTeX block.
