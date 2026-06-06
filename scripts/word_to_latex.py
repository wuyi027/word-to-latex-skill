#!/usr/bin/env python3
"""
Convert a Word .docx document into a compile-ready LaTeX project.

Core features:
- Preserves document order for paragraphs, headings, tables, and inline images.
- Converts common Word OMML equations to LaTeX math.
- Extracts embedded images to figures/ and emits figure environments.
- Rebuilds Word tables as booktabs/longtable LaTeX tables.
- Detects references/bibliography sections and emits references.tex + references.bib.
- Uses a LaTeX template with %%TITLE%%, %%CONTENT%%, %%BIBLIOGRAPHY%% placeholders.
- Optionally compiles and fails non-zero on LaTeX errors.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import zipfile
import shlex
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

try:
    from lxml import etree
except ImportError:  # Keep --check-deps usable even before dependencies are installed.
    etree = None

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}

LATEX_SPECIALS = {
    "\\": r"\textbackslash{}",
    "&": r"\&",
    "%": r"\%",
    "$": r"\$",
    "#": r"\#",
    "_": r"\_",
    "{": r"\{",
    "}": r"\}",
    "~": r"\textasciitilde{}",
    "^": r"\textasciicircum{}",
}

REFERENCE_HEADING_RE = re.compile(r"^(references|bibliography|参考文献|参考资料|文献)$", re.I)
FIG_CAPTION_RE = re.compile(r"^(fig(?:ure)?\.?\s*\d+|图\s*\d+|figure\s*\d+)", re.I)
TABLE_CAPTION_RE = re.compile(r"^(table\s*\d+|表\s*\d+)", re.I)
DOI_RE = re.compile(r"10\.\d{4,9}/[-._;()/:A-Z0-9]+", re.I)
URL_RE = re.compile(r"https?://\S+", re.I)
MATH_SPAN_RE = re.compile(r"(\\\[[\s\S]*?\\\]|\$[^$\n]+\$)")
IMAGE_MARKER_RE = re.compile(r"\[\[IMAGE:([^\]]+)\]\]")
NUMBERED_HEADING_RE = re.compile(r"^(\d+(?:\.\d+){0,3})\.?\s+(.+)$")

OMML_SYMBOLS = {
    "±": r"\pm",
    "×": r"\times",
    "÷": r"\div",
    "≤": r"\leq",
    "≥": r"\geq",
    "≠": r"\neq",
    "≈": r"\approx",
    "∞": r"\infty",
    "∂": r"\partial",
    "∇": r"\nabla",
    "∈": r"\in",
    "∉": r"\notin",
    "⊂": r"\subset",
    "⊆": r"\subseteq",
    "∑": r"\sum",
    "∏": r"\prod",
    "∫": r"\int",
    "√": r"\sqrt",
    "→": r"\to",
    "←": r"\leftarrow",
    "↔": r"\leftrightarrow",
    "α": r"\alpha",
    "β": r"\beta",
    "γ": r"\gamma",
    "δ": r"\delta",
    "ε": r"\epsilon",
    "θ": r"\theta",
    "λ": r"\lambda",
    "μ": r"\mu",
    "π": r"\pi",
    "σ": r"\sigma",
    "φ": r"\phi",
    "ω": r"\omega",
    "Γ": r"\Gamma",
    "Δ": r"\Delta",
    "Θ": r"\Theta",
    "Λ": r"\Lambda",
    "Π": r"\Pi",
    "Σ": r"\Sigma",
    "Φ": r"\Phi",
    "Ω": r"\Omega",
}

PYTHON_DEPENDENCIES = {
    "lxml": "lxml>=5.0.0",
    "docx": "python-docx>=1.1.0",
    "pylatexenc": "pylatexenc>=2.10",
}
LATEX_ENGINES = ("xelatex", "pdflatex", "lualatex")
OPTIONAL_TOOLS = ("pandoc",)


def dependency_report(engine: str = "xelatex") -> Dict[str, object]:
    """Return a machine-readable dependency report without requiring lxml to import."""
    import importlib.util
    python_version_ok = sys.version_info >= (3, 9)
    python_packages = {
        module: {
            "installed": importlib.util.find_spec(module) is not None,
            "install_spec": spec,
        }
        for module, spec in PYTHON_DEPENDENCIES.items()
    }
    commands = {
        "latexmk": shutil.which("latexmk") is not None,
        "xelatex": shutil.which("xelatex") is not None,
        "pdflatex": shutil.which("pdflatex") is not None,
        "lualatex": shutil.which("lualatex") is not None,
        "pandoc": shutil.which("pandoc") is not None,
    }
    requested_engine_available = commands.get(engine, False)
    any_latex_engine_available = any(commands[e] for e in LATEX_ENGINES)
    missing_python = [spec for module, spec in PYTHON_DEPENDENCIES.items() if not python_packages[module]["installed"]]
    return {
        "python": {
            "version": f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}",
            "requires": ">=3.9",
            "ok": python_version_ok,
        },
        "python_packages": python_packages,
        "commands": commands,
        "requested_engine": engine,
        "requested_engine_available": requested_engine_available,
        "any_latex_engine_available": any_latex_engine_available,
        "missing_python_packages": missing_python,
        "missing_required_commands": [] if any_latex_engine_available else ["xelatex|pdflatex|lualatex"],
        "missing_optional_commands": [tool for tool in OPTIONAL_TOOLS if not commands.get(tool, False)],
        "ok_for_conversion": python_version_ok and not missing_python,
        "ok_for_compilation": any_latex_engine_available,
    }


def print_dependency_report(report: Dict[str, object]) -> None:
    print("Dependency check")
    print("================")
    py = report["python"]
    print(f"Python: {py['version']} (requires {py['requires']}) -> {'OK' if py['ok'] else 'MISSING/OLD'}")
    print("\nPython packages:")
    for module, info in report["python_packages"].items():
        print(f"- {info['install_spec']}: {'OK' if info['installed'] else 'MISSING'}")
    print("\nExternal commands:")
    for cmd, installed in report["commands"].items():
        kind = "optional" if cmd in OPTIONAL_TOOLS else "required for compilation"
        print(f"- {cmd}: {'OK' if installed else 'MISSING'} ({kind})")
    if report["missing_python_packages"]:
        print("\nMissing Python packages can be installed with:")
        print("  " + " ".join(shlex.quote(x) for x in [sys.executable, "-m", "pip", "install", *report["missing_python_packages"]]))
    if not report["any_latex_engine_available"]:
        print("\nNo LaTeX engine was found. Install TeX Live, MacTeX, or MiKTeX, then make xelatex/pdflatex/lualatex available in PATH.")
    if report["missing_optional_commands"]:
        print("\nOptional tools missing: " + ", ".join(report["missing_optional_commands"]))
        print("Pandoc is optional and only needed for advanced fallback conversions.")


def maybe_install_missing_python_packages(report: Dict[str, object], assume_yes: bool = False) -> bool:
    missing = list(report.get("missing_python_packages", []))
    if not missing:
        return True
    cmd = [sys.executable, "-m", "pip", "install", *missing]
    if not assume_yes:
        if not sys.stdin.isatty():
            print("\nMissing Python packages were found, but this is not an interactive terminal.", file=sys.stderr)
            print("Run with --install-missing --yes to allow pip installation, or install manually:", file=sys.stderr)
            print("  " + " ".join(shlex.quote(x) for x in cmd), file=sys.stderr)
            return False
        answer = input("\nInstall missing Python packages now with pip? This may require network access and write permission. [y/N] ").strip().lower()
        if answer not in {"y", "yes"}:
            print("Installation skipped. Install manually before conversion:")
            print("  " + " ".join(shlex.quote(x) for x in cmd))
            return False
    print("Installing missing Python packages:")
    print("  " + " ".join(shlex.quote(x) for x in cmd))
    proc = subprocess.run(cmd)
    return proc.returncode == 0


def ensure_runtime_dependencies() -> None:
    if sys.version_info < (3, 9):
        raise RuntimeError("Python >= 3.9 is required.")
    if etree is None:
        raise RuntimeError("Missing dependency: lxml. Run `python -m pip install -r requirements.txt` or `python scripts/word_to_latex.py --check-deps --install-missing`.")


@dataclass
class ParagraphItem:
    text: str
    style: str = ""
    level: Optional[int] = None
    images: List[str] = field(default_factory=list)


@dataclass
class TableItem:
    rows: List[List[str]]


DocumentItem = ParagraphItem | TableItem


def sanitize_filename(name: str) -> str:
    name = re.sub(r"[^A-Za-z0-9_.-]+", "_", name.strip())
    return name.strip("._") or "file"


def latex_escape(text: str) -> str:
    if text is None:
        return ""
    text = text.replace("\u00a0", " ").replace("\r", " ")
    return "".join(LATEX_SPECIALS.get(ch, ch) for ch in text)


def latex_escape_preserving_math(text: str) -> str:
    if not text:
        return ""
    parts: List[str] = []
    last = 0
    for match in MATH_SPAN_RE.finditer(text):
        parts.append(latex_escape(text[last:match.start()]))
        parts.append(match.group(0))
        last = match.end()
    parts.append(latex_escape(text[last:]))
    return "".join(parts)


def normalize_ws(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def omml_arg(node: etree._Element, name: str) -> str:
    children = node.xpath(f"./m:{name}/*", namespaces=NS)
    return omml_children_to_latex(children)


def omml_text_to_latex(text: str) -> str:
    out: List[str] = []
    for ch in text or "":
        if ch.isspace():
            out.append(" ")
        elif ch in OMML_SYMBOLS:
            out.append(OMML_SYMBOLS[ch])
        elif ch in {"\\", "{", "}", "&", "%", "$", "#", "_", "^", "~"}:
            out.append(LATEX_SPECIALS[ch])
        else:
            out.append(ch)
    return "".join(out)


def omml_element_to_latex(node: etree._Element) -> str:
    tag = etree.QName(node).localname
    if tag in {"oMath", "oMathPara", "e", "num", "den", "sub", "sup", "deg", "rad", "naryPr", "lim", "limLoc"}:
        return omml_children_to_latex(node)
    if tag == "r":
        return "".join(omml_text_to_latex(t) for t in node.xpath(".//m:t/text()", namespaces=NS))
    if tag == "t":
        return omml_text_to_latex(node.text or "")
    if tag == "f":
        return r"\frac{%s}{%s}" % (omml_arg(node, "num"), omml_arg(node, "den"))
    if tag == "sSub":
        return "{%s}_{%s}" % (omml_arg(node, "e"), omml_arg(node, "sub"))
    if tag == "sSup":
        return "{%s}^{%s}" % (omml_arg(node, "e"), omml_arg(node, "sup"))
    if tag == "sSubSup":
        return "{%s}_{%s}^{%s}" % (omml_arg(node, "e"), omml_arg(node, "sub"), omml_arg(node, "sup"))
    if tag == "rad":
        deg = omml_arg(node, "deg")
        rad = omml_arg(node, "e")
        return r"\sqrt[%s]{%s}" % (deg, rad) if deg else r"\sqrt{%s}" % rad
    if tag == "d":
        content = omml_arg(node, "e")
        beg = "".join(node.xpath("./m:dPr/m:begChr/@m:val", namespaces=NS)) or "("
        end = "".join(node.xpath("./m:dPr/m:endChr/@m:val", namespaces=NS)) or ")"
        return r"\left%s %s \right%s" % (beg, content, end)
    if tag == "nary":
        op = "".join(node.xpath("./m:naryPr/m:chr/@m:val", namespaces=NS)) or "∑"
        op_latex = OMML_SYMBOLS.get(op, omml_text_to_latex(op))
        sub = omml_arg(node, "sub")
        sup = omml_arg(node, "sup")
        base = op_latex
        if sub:
            base += "_{%s}" % sub
        if sup:
            base += "^{%s}" % sup
        return base + " " + omml_arg(node, "e")
    if tag == "limLow":
        return r"\underset{%s}{%s}" % (omml_arg(node, "lim"), omml_arg(node, "e"))
    if tag == "limUpp":
        return r"\overset{%s}{%s}" % (omml_arg(node, "lim"), omml_arg(node, "e"))
    if tag == "bar":
        return r"\overline{%s}" % omml_arg(node, "e")
    if tag == "acc":
        accent = "".join(node.xpath("./m:accPr/m:chr/@m:val", namespaces=NS)) or "^"
        expr = omml_arg(node, "e")
        accents = {"^": r"\hat", "~": r"\tilde", "¯": r"\bar", "→": r"\vec", ".": r"\dot", "¨": r"\ddot"}
        return "%s{%s}" % (accents.get(accent, r"\hat"), expr)
    if tag == "m":
        rows = []
        for mr in node.xpath("./m:mr", namespaces=NS):
            cells = [omml_children_to_latex(e) for e in mr.xpath("./m:e", namespaces=NS)]
            rows.append(" & ".join(cells))
        return r"\begin{matrix}%s\end{matrix}" % (r" \\ ".join(rows))
    if tag in {"box", "groupChr"}:
        return omml_arg(node, "e")
    return omml_children_to_latex(node)


def omml_children_to_latex(node_or_nodes: etree._Element | Sequence[etree._Element]) -> str:
    nodes = list(node_or_nodes) if not isinstance(node_or_nodes, etree._Element) else list(node_or_nodes)
    return normalize_ws("".join(omml_element_to_latex(child) for child in nodes))


def omml_to_latex_math(node: etree._Element) -> str:
    latex = omml_element_to_latex(node)
    if not latex:
        return ""
    tag = etree.QName(node).localname
    if tag == "oMathPara":
        return r"\[%s\]" % latex
    return "$%s$" % latex


def parse_xml_from_docx(docx_path: Path, internal_path: str) -> etree._Element:
    ensure_runtime_dependencies()
    with zipfile.ZipFile(docx_path) as zf:
        with zf.open(internal_path) as f:
            return etree.parse(f).getroot()


def load_relationships(docx_path: Path) -> Dict[str, str]:
    rel_path = "word/_rels/document.xml.rels"
    rels: Dict[str, str] = {}
    try:
        root = parse_xml_from_docx(docx_path, rel_path)
    except KeyError:
        return rels
    rel_ns = {"pr": "http://schemas.openxmlformats.org/package/2006/relationships"}
    for rel in root.xpath("//pr:Relationship", namespaces=rel_ns):
        rid = rel.get("Id")
        target = rel.get("Target")
        if rid and target:
            rels[rid] = target
    return rels


def extract_media(docx_path: Path, output_figures: Path) -> Dict[str, str]:
    """Extract word/media/* into output_figures and return internal path -> local filename."""
    output_figures.mkdir(parents=True, exist_ok=True)
    media_map: Dict[str, str] = {}
    counter = 0
    with zipfile.ZipFile(docx_path) as zf:
        for name in zf.namelist():
            if not name.startswith("word/media/"):
                continue
            suffix = Path(name).suffix.lower() or ".bin"
            counter += 1
            local_name = f"figure_{counter:03d}{suffix}"
            local_path = output_figures / local_name
            with zf.open(name) as src, local_path.open("wb") as dst:
                shutil.copyfileobj(src, dst)
            media_map[name] = local_name
            media_map[name.replace("word/", "")] = local_name
            media_map[Path(name).name] = local_name
    return media_map


def paragraph_text(p: etree._Element) -> str:
    pieces: List[str] = []

    def visit(node: etree._Element) -> None:
        tag = etree.QName(node).localname
        namespace = etree.QName(node).namespace
        if namespace == NS["m"] and tag in {"oMath", "oMathPara"}:
            pieces.append(omml_to_latex_math(node))
            return
        if namespace == NS["w"] and tag == "t":
            pieces.append(node.text or "")
            return
        if namespace == NS["w"] and tag == "tab":
            pieces.append("\t")
            return
        if namespace == NS["w"] and tag == "br":
            pieces.append(" ")
            return
        for child in node:
            visit(child)

    visit(p)
    return normalize_ws("".join(pieces))


def paragraph_style_and_level(p: etree._Element) -> Tuple[str, Optional[int]]:
    style_nodes = p.xpath("./w:pPr/w:pStyle/@w:val", namespaces=NS)
    style = style_nodes[0] if style_nodes else ""
    lower = style.lower()
    # Word built-in styles often appear as Heading1, Heading2; localized docs may still keep style IDs.
    m = re.search(r"heading\s*(\d+)|标题\s*(\d+)", lower, re.I)
    if m:
        return style, int(m.group(1) or m.group(2))
    # Also infer from outline level if present.
    outline = p.xpath("./w:pPr/w:outlineLvl/@w:val", namespaces=NS)
    if outline:
        try:
            return style, int(outline[0]) + 1
        except ValueError:
            pass
    return style, None


def image_targets_from_element(element: etree._Element, rels: Dict[str, str], media_map: Dict[str, str]) -> List[str]:
    image_files: List[str] = []
    rids = element.xpath(".//a:blip/@r:embed", namespaces=NS)
    for rid in rids:
        target = rels.get(rid, "")
        candidates = []
        if target:
            candidates.extend([target, f"word/{target}", Path(target).name])
        local = next((media_map[c] for c in candidates if c in media_map), None)
        if local and local not in image_files:
            image_files.append(local)
    return image_files


def paragraph_image_targets(p: etree._Element, rels: Dict[str, str], media_map: Dict[str, str]) -> List[str]:
    return image_targets_from_element(p, rels, media_map)


def table_rows(tbl: etree._Element, rels: Dict[str, str], media_map: Dict[str, str]) -> List[List[str]]:
    rows: List[List[str]] = []
    for tr in tbl.xpath("./w:tr", namespaces=NS):
        row: List[str] = []
        for tc in tr.xpath("./w:tc", namespaces=NS):
            texts = [paragraph_text(p) for p in tc.xpath("./w:p", namespaces=NS)]
            images = image_targets_from_element(tc, rels, media_map)
            image_markers = [f"[[IMAGE:{img}]]" for img in images]
            row.append(normalize_ws(" ".join(t for t in [*texts, *image_markers] if t)))
        if any(cell for cell in row):
            rows.append(row)
    if not rows:
        return rows
    max_cols = max(len(r) for r in rows)
    return [r + [""] * (max_cols - len(r)) for r in rows]


def parse_docx(docx_path: Path, figures_dir: Path) -> Tuple[List[DocumentItem], str]:
    rels = load_relationships(docx_path)
    media_map = extract_media(docx_path, figures_dir)
    root = parse_xml_from_docx(docx_path, "word/document.xml")
    body = root.find("w:body", namespaces=NS)
    items: List[DocumentItem] = []
    if body is None:
        return items, docx_path.stem

    for child in body:
        tag = etree.QName(child).localname
        if tag == "p":
            text = paragraph_text(child)
            style, level = paragraph_style_and_level(child)
            images = paragraph_image_targets(child, rels, media_map)
            if text or images:
                items.append(ParagraphItem(text=text, style=style, level=level, images=images))
        elif tag == "tbl":
            rows = table_rows(child, rels, media_map)
            if rows:
                items.append(TableItem(rows=rows))

    title = infer_title(items, docx_path.stem)
    return items, title


def infer_title(items: Sequence[DocumentItem], fallback: str) -> str:
    for item in items[:20]:
        if isinstance(item, ParagraphItem) and item.text:
            if item.level == 1 or "title" in item.style.lower() or "标题" in item.style:
                return item.text
    for item in items[:10]:
        if isinstance(item, ParagraphItem) and item.text and len(item.text) < 180:
            return item.text
    return fallback


def heading_command(level: int) -> str:
    if level <= 1:
        return "section"
    if level == 2:
        return "subsection"
    if level == 3:
        return "subsubsection"
    return "paragraph"


def numbered_heading_command(text: str) -> Optional[Tuple[str, str]]:
    match = NUMBERED_HEADING_RE.match(normalize_ws(text))
    if not match:
        return None
    number, title = match.groups()
    if len(title) > 90 or re.match(r"^(we|our|this|the)\b", title, re.I):
        return None
    depth = number.count(".") + 1
    if depth == 1:
        return "section", title
    if depth == 2:
        return "subsection", title
    if depth == 3:
        return "subsubsection", title
    return "paragraph", title


def extract_references(items: Sequence[DocumentItem]) -> Tuple[List[DocumentItem], List[str]]:
    body: List[DocumentItem] = []
    refs: List[str] = []
    in_refs = False
    for item in items:
        if isinstance(item, ParagraphItem):
            txt = normalize_ws(item.text)
            if txt and REFERENCE_HEADING_RE.match(txt):
                in_refs = True
                continue
            if in_refs:
                if txt:
                    refs.append(txt)
                continue
        if not in_refs:
            body.append(item)
    return body, refs


def split_caption(text: str, kind: str) -> str:
    if not text:
        return ""
    pat = FIG_CAPTION_RE if kind == "figure" else TABLE_CAPTION_RE
    return normalize_ws(pat.sub("", text).strip(" .:：-—")) or text


def latex_table_cell(cell: str, cols: int) -> str:
    if not cell:
        return ""
    width = max(0.10, min(0.18, 0.9 / max(cols, 1)))
    parts: List[str] = []
    last = 0
    for match in IMAGE_MARKER_RE.finditer(cell):
        prefix = cell[last:match.start()].strip()
        if prefix:
            parts.append(latex_escape_preserving_math(prefix))
        img = match.group(1)
        parts.append(
            "\\begin{minipage}{%.3f\\textwidth}\\centering"
            "\\includegraphics[width=\\linewidth]{figures/%s}"
            "\\end{minipage}" % (width, img)
        )
        last = match.end()
    suffix = cell[last:].strip()
    if suffix:
        parts.append(latex_escape_preserving_math(suffix))
    if not parts:
        return latex_escape_preserving_math(cell)
    return r" \\ ".join(parts)


def table_to_latex(rows: List[List[str]], caption: str, label: str, use_longtable: bool = False) -> str:
    if not rows:
        return ""
    cols = max(len(r) for r in rows)
    escaped = [[latex_table_cell(c, cols) for c in r] for r in rows]
    align = "@{}" + "l" * cols + "@{}"
    lines: List[str] = []
    if use_longtable or len(rows) > 25:
        lines.append(f"\\begin{{longtable}}{{{align}}}")
        if caption:
            lines.append(f"\\caption{{{latex_escape_preserving_math(caption)}}}\\label{{{label}}}\\\\")
        lines.append("\\toprule")
        lines.append(" & ".join(escaped[0]) + r" \\")
        lines.append("\\midrule")
        lines.append("\\endfirsthead")
        lines.append("\\toprule")
        lines.append(" & ".join(escaped[0]) + r" \\")
        lines.append("\\midrule")
        lines.append("\\endhead")
        for r in escaped[1:]:
            lines.append(" & ".join(r) + r" \\")
        lines.append("\\bottomrule")
        lines.append("\\end{longtable}")
    else:
        lines.append("\\begin{table}[htbp]")
        lines.append("\\centering")
        if caption:
            lines.append(f"\\caption{{{latex_escape_preserving_math(caption)}}}")
            lines.append(f"\\label{{{label}}}")
        lines.append(f"\\begin{{tabular}}{{{align}}}")
        lines.append("\\toprule")
        lines.append(" & ".join(escaped[0]) + r" \\")
        lines.append("\\midrule")
        for r in escaped[1:]:
            lines.append(" & ".join(r) + r" \\")
        lines.append("\\bottomrule")
        lines.append("\\end{tabular}")
        lines.append("\\end{table}")
    return "\n".join(lines) + "\n"


def items_to_latex(items: Sequence[DocumentItem]) -> str:
    out: List[str] = []
    i = 0
    fig_no = 0
    tbl_no = 0
    while i < len(items):
        item = items[i]
        if isinstance(item, ParagraphItem):
            if item.images:
                caption = ""
                if i + 1 < len(items) and isinstance(items[i + 1], ParagraphItem):
                    nxt = items[i + 1]
                    if FIG_CAPTION_RE.match(nxt.text):
                        caption = split_caption(nxt.text, "figure")
                        i += 1
                if item.text and not caption and not FIG_CAPTION_RE.match(item.text):
                    out.append(latex_escape_preserving_math(item.text) + "\n")
                for img in item.images:
                    fig_no += 1
                    if caption:
                        out.append("\\begin{figure}[htbp]")
                        out.append("\\centering")
                        out.append(f"\\includegraphics[width=0.85\\textwidth]{{figures/{img}}}")
                        out.append(f"\\caption{{{latex_escape_preserving_math(caption)}}}")
                        out.append(f"\\label{{fig:{fig_no}}}")
                        out.append("\\end{figure}\n")
                    else:
                        out.append("\\begin{center}")
                        out.append(f"\\includegraphics[width=0.72\\textwidth]{{figures/{img}}}")
                        out.append("\\end{center}\n")
            elif item.text:
                if item.level:
                    out.append(f"\\{heading_command(item.level)}{{{latex_escape_preserving_math(item.text)}}}\n")
                elif numbered_heading_command(item.text):
                    command, heading = numbered_heading_command(item.text) or ("paragraph", item.text)
                    out.append(f"\\{command}{{{latex_escape_preserving_math(heading)}}}\n")
                elif FIG_CAPTION_RE.match(item.text):
                    # Standalone orphan captions are kept as text, but marked.
                    out.append("% Orphan figure caption from Word\n" + latex_escape_preserving_math(item.text) + "\n")
                elif TABLE_CAPTION_RE.match(item.text):
                    out.append("% Orphan table caption from Word\n" + latex_escape_preserving_math(item.text) + "\n")
                else:
                    out.append(latex_escape_preserving_math(item.text) + "\n")
        else:
            caption = ""
            # Prefer previous table caption if immediately before; otherwise next caption.
            if out and False:
                pass
            if i + 1 < len(items) and isinstance(items[i + 1], ParagraphItem):
                nxt = items[i + 1]
                if TABLE_CAPTION_RE.match(nxt.text):
                    caption = split_caption(nxt.text, "table")
                    i += 1
            tbl_no += 1
            out.append(table_to_latex(item.rows, caption or f"Table {tbl_no}", f"tab:{tbl_no}"))
        i += 1
    return "\n".join(out).strip() + "\n"


def references_to_tex(refs: Sequence[str]) -> str:
    if not refs:
        return ""
    lines = ["\\begin{thebibliography}{99}"]
    for idx, ref in enumerate(refs, start=1):
        lines.append(f"\\bibitem{{ref{idx}}} {latex_escape_preserving_math(ref)}")
    lines.append("\\end{thebibliography}")
    return "\n".join(lines) + "\n"


def references_to_bib(refs: Sequence[str]) -> str:
    entries: List[str] = []
    for idx, ref in enumerate(refs, start=1):
        doi = DOI_RE.search(ref)
        url = URL_RE.search(ref)
        fields = [f"  note = {{{ref}}}"]
        if doi:
            fields.append(f"  doi = {{{doi.group(0).rstrip('.,;')}}}")
        if url:
            fields.append(f"  url = {{{url.group(0).rstrip('.,;')}}}")
        entries.append("@misc{ref%d,\n%s\n}" % (idx, ",\n".join(fields)))
    return "\n\n".join(entries) + ("\n" if entries else "")


def load_template(template: Optional[Path]) -> str:
    if template and template.exists():
        return template.read_text(encoding="utf-8")
    default_template = Path(__file__).resolve().parents[1] / "assets" / "default_template.tex"
    return default_template.read_text(encoding="utf-8")


def prepare_template_content(template_text: str, title: str, content: str) -> Tuple[str, str]:
    """Apply light journal-template cleanup without changing the user's source text."""
    if "elsarticle" not in template_text or "\\begin{frontmatter}" not in template_text:
        return template_text, content

    lines = content.splitlines()
    if lines and normalize_ws(lines[0]) == normalize_ws(title):
        lines = lines[1:]
    while lines and not lines[0].strip():
        lines = lines[1:]

    abstract = ""
    keywords = ""
    if lines and lines[0].lower().startswith("abstract:"):
        abstract = lines[0].split(":", 1)[1].strip()
        lines = lines[1:]
    while lines and not lines[0].strip():
        lines = lines[1:]
    if lines and lines[0].lower().startswith("keywords:"):
        keywords = lines[0].split(":", 1)[1].strip()
        lines = lines[1:]

    frontmatter_bits: List[str] = []
    if abstract and "\\begin{abstract}" not in template_text:
        frontmatter_bits.append("\\begin{abstract}\n" + latex_escape_preserving_math(abstract) + "\n\\end{abstract}")
    if keywords and "\\begin{keyword}" not in template_text:
        keyword_text = " \\sep ".join(latex_escape_preserving_math(k.strip()) for k in keywords.split(",") if k.strip())
        if keyword_text:
            frontmatter_bits.append("\\begin{keyword}\n" + keyword_text + "\n\\end{keyword}")
    if frontmatter_bits:
        template_text = template_text.replace("\\end{frontmatter}", "\n\n".join(frontmatter_bits) + "\n\\end{frontmatter}")

    if "\\setlength{\\parindent}" not in template_text:
        template_text = template_text.replace("\\usepackage{xeCJK}\n", "\\usepackage{xeCJK}\n\\setlength{\\parindent}{0pt}\n\\setlength{\\parskip}{0.35em}\n")
    return template_text, "\n".join(lines).lstrip() + "\n"


def rasterize_vector_graphics(output_dir: Path, main_tex: str) -> Tuple[str, int]:
    """Convert referenced EMF/WMF graphics to uniquely named PNG files when ImageMagick is available."""
    refs = sorted(set(re.findall(r"figures/([^{}]+\.(?:emf|wmf))", main_tex, flags=re.I)))
    if not refs:
        return main_tex, 0
    magick = shutil.which("magick")
    if not magick:
        return main_tex, 0
    converted = 0
    figures_dir = output_dir / "figures"
    for ref in refs:
        src = figures_dir / ref
        if not src.exists():
            continue
        dst = src.with_name(src.stem + "_raster.png")
        try:
            subprocess.run([magick, str(src), str(dst)], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        except Exception:
            continue
        main_tex = main_tex.replace(f"figures/{ref}", f"figures/{dst.name}")
        converted += 1
    return main_tex, converted


def preprocess_mathtype_equations(docx_path: Path, output_dir: Path, visible: bool = False, wait_seconds: float = 0.7) -> Path:
    """Ask Word/MathType to convert selected MathType equations to LaTeX in a copied DOCX.

    This uses Windows COM + SendKeys because MathType's Alt+\ conversion is a Word UI command,
    not data exposed directly in DOCX XML. The original file is never modified.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    preprocessed = output_dir / f"{docx_path.stem}_mathtype_preprocessed.docx"
    script = r"""
param(
  [string]$InputDocx,
  [string]$OutputDocx,
  [string]$Visible,
  [double]$WaitSeconds
)
$ErrorActionPreference = 'Stop'
$wdFormatDocumentDefault = 16
$wdInlineShapeEmbeddedOLEObject = 1
$wdInlineShapeLinkedOLEObject = 2
$word = $null
$doc = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = [System.Convert]::ToBoolean($Visible)
  $word.DisplayAlerts = 0
  $doc = $word.Documents.Open($InputDocx)
  $shell = New-Object -ComObject WScript.Shell
  $converted = 0

  foreach ($shape in @($doc.InlineShapes)) {
    if ($shape.Type -eq $wdInlineShapeEmbeddedOLEObject -or $shape.Type -eq $wdInlineShapeLinkedOLEObject) {
      $shape.Select()
      Start-Sleep -Seconds $WaitSeconds
      $shell.SendKeys('%\')
      Start-Sleep -Seconds $WaitSeconds
      $converted += 1
    }
  }

  foreach ($shape in @($doc.Shapes)) {
    try {
      $shape.Select()
      Start-Sleep -Seconds $WaitSeconds
      $shell.SendKeys('%\')
      Start-Sleep -Seconds $WaitSeconds
      $converted += 1
    } catch {
      # Floating shapes that cannot be selected are skipped.
    }
  }

  $doc.SaveAs2($OutputDocx, $wdFormatDocumentDefault)
  Write-Output $converted
} finally {
  if ($doc -ne $null) { $doc.Close($false) | Out-Null }
  if ($word -ne $null) { $word.Quit() | Out-Null }
}
"""
    script_path = output_dir / "preprocess_mathtype.ps1"
    script_path.write_text(script, encoding="utf-8")
    cmd = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script_path.resolve()),
        "-InputDocx",
        str(docx_path.resolve()),
        "-OutputDocx",
        str(preprocessed.resolve()),
        "-Visible",
        "true" if visible else "false",
        "-WaitSeconds",
        str(wait_seconds),
    ]
    proc = subprocess.run(cmd, text=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if proc.returncode != 0:
        raise RuntimeError(
            "MathType preprocessing failed. It requires Microsoft Word, MathType, and an interactive Windows session.\n"
            + proc.stderr.strip()
        )
    return preprocessed


def render_project(docx_path: Path, output_dir: Path, template: Optional[Path]) -> Dict[str, object]:
    output_dir.mkdir(parents=True, exist_ok=True)
    figures_dir = output_dir / "figures"
    items, title = parse_docx(docx_path, figures_dir)
    body_items, refs = extract_references(items)
    content = items_to_latex(body_items)

    refs_tex = references_to_tex(refs)
    refs_bib = references_to_bib(refs)
    (output_dir / "references.tex").write_text(refs_tex, encoding="utf-8")
    (output_dir / "references.bib").write_text(refs_bib, encoding="utf-8")

    bibliography_block = "\\input{references.tex}" if refs else "% No references detected"
    template_text = load_template(template)
    template_text, content = prepare_template_content(template_text, title, content)
    main_tex = (template_text
                .replace("%%TITLE%%", latex_escape_preserving_math(title))
                .replace("%%CONTENT%%", content)
                .replace("%%BIBLIOGRAPHY%%", bibliography_block))
    main_tex, rasterized = rasterize_vector_graphics(output_dir, main_tex)
    main_path = output_dir / "main.tex"
    main_path.write_text(main_tex, encoding="utf-8")
    return {
        "title": title,
        "tex": str(main_path),
        "figures": len(list(figures_dir.glob("*"))) if figures_dir.exists() else 0,
        "rasterized_vector_graphics": rasterized,
        "tables": sum(isinstance(x, TableItem) for x in body_items),
        "references": len(refs),
    }


def find_executable(candidates: Sequence[str]) -> Optional[str]:
    for c in candidates:
        p = shutil.which(c)
        if p:
            return p
    return None


def compile_latex(output_dir: Path, engine: str) -> Dict[str, object]:
    main = output_dir / "main.tex"
    if not main.exists():
        raise FileNotFoundError(main)
    log_path = output_dir / "compile.log"
    latexmk = shutil.which("latexmk")
    commands: List[List[str]] = []
    if latexmk:
        commands.append([latexmk, f"-{engine}", "-interaction=nonstopmode", "-halt-on-error", "main.tex"])
    else:
        exe = find_executable([engine, "xelatex", "pdflatex"])
        if not exe:
            return {"compiled": False, "reason": "No LaTeX compiler found in PATH."}
        commands.append([exe, "-interaction=nonstopmode", "-halt-on-error", "main.tex"])
        commands.append([exe, "-interaction=nonstopmode", "-halt-on-error", "main.tex"])

    full_log = []
    ok = True
    for cmd in commands:
        proc = subprocess.run(cmd, cwd=output_dir, text=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        full_log.append("$ " + " ".join(cmd) + "\n" + proc.stdout)
        if proc.returncode != 0:
            ok = False
            break
    log_path.write_text("\n\n".join(full_log), encoding="utf-8", errors="ignore")
    return {"compiled": ok, "pdf": str(output_dir / "main.pdf"), "log": str(log_path)}


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Convert DOCX to a compile-ready LaTeX project.")
    parser.add_argument("input", type=Path, nargs="?", help="Input .docx file")
    parser.add_argument("-o", "--output", type=Path, default=Path("latex_project"), help="Output LaTeX project directory")
    parser.add_argument("-t", "--template", type=Path, default=None, help="Optional .tex template with %%CONTENT%% placeholder")
    parser.add_argument("--engine", default="xelatex", choices=["xelatex", "pdflatex", "lualatex"], help="LaTeX engine")
    parser.add_argument("--compile", action="store_true", help="Compile main.tex after conversion")
    parser.add_argument("--preprocess-mathtype", action="store_true", help="Use Word/MathType UI automation to convert MathType equations to LaTeX in a copied DOCX before conversion")
    parser.add_argument("--mathtype-visible", action="store_true", help="Show Word while --preprocess-mathtype runs")
    parser.add_argument("--mathtype-wait", type=float, default=0.7, help="Seconds to wait after each MathType Alt+\\ command")
    parser.add_argument("--json", action="store_true", help="Print machine-readable JSON summary")
    parser.add_argument("--check-deps", action="store_true", help="Check Python packages and LaTeX tools, then exit unless an input file is also provided")
    parser.add_argument("--install-missing", action="store_true", help="Prompt for permission to install missing Python packages with pip during --check-deps")
    parser.add_argument("--yes", action="store_true", help="Assume yes for --install-missing in non-interactive environments")
    args = parser.parse_args(argv)

    if args.check_deps:
        report = dependency_report(args.engine)
        if args.json:
            print(json.dumps(report, ensure_ascii=False, indent=2))
        else:
            print_dependency_report(report)
        if args.install_missing and report.get("missing_python_packages"):
            ok = maybe_install_missing_python_packages(report, assume_yes=args.yes)
            if not ok:
                return 1
            report = dependency_report(args.engine)
            if args.json:
                print(json.dumps({"after_install": report}, ensure_ascii=False, indent=2))
            else:
                print("\nAfter installation:")
                print_dependency_report(report)
        if args.input is None:
            return 0 if dependency_report(args.engine).get("ok_for_conversion") else 1

    if args.input is None:
        parser.error("the following argument is required unless --check-deps is used: input")

    ensure_runtime_dependencies()

    if args.input.suffix.lower() != ".docx":
        print("Only .docx input is supported. Convert .doc to .docx first.", file=sys.stderr)
        return 2
    if not args.input.exists():
        print(f"Input file not found: {args.input}", file=sys.stderr)
        return 2

    input_docx = args.input
    preprocessing_summary: Optional[Dict[str, object]] = None
    if args.preprocess_mathtype:
        try:
            input_docx = preprocess_mathtype_equations(args.input, args.output, visible=args.mathtype_visible, wait_seconds=args.mathtype_wait)
            preprocessing_summary = {
                "mathtype": True,
                "input": str(args.input),
                "preprocessed_docx": str(input_docx),
            }
        except Exception as exc:
            print(str(exc), file=sys.stderr)
            return 1

    summary = render_project(input_docx, args.output, args.template)
    if preprocessing_summary:
        summary["preprocessing"] = preprocessing_summary
    if args.compile:
        summary["compile"] = compile_latex(args.output, args.engine)
        if isinstance(summary["compile"], dict) and not summary["compile"].get("compiled"):
            if args.json:
                print(json.dumps(summary, ensure_ascii=False, indent=2))
            else:
                print(f"Converted but compilation failed or unavailable. See: {summary['compile'].get('log') or summary['compile'].get('reason')}")
            return 1

    if args.json:
        print(json.dumps(summary, ensure_ascii=False, indent=2))
    else:
        print(f"LaTeX project written to: {args.output}")
        print(f"main.tex: {summary['tex']}")
        print(f"figures={summary['figures']} tables={summary['tables']} references={summary['references']}")
        if args.compile:
            print(f"PDF: {summary['compile'].get('pdf')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
