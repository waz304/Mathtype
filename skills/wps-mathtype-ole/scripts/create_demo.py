"""Generate a synthetic DOCX matching examples/manifest.json."""

import argparse
from pathlib import Path

from docx import Document
from docx.shared import Pt


def main():
    parser = argparse.ArgumentParser(description="生成带有示例公式占位符的 DOCX")
    default_output = Path(__file__).resolve().parents[1] / "examples" / "demo_source.docx"
    parser.add_argument("--output", type=Path, default=default_output)
    args = parser.parse_args()
    output = args.output.resolve()
    if output.suffix.lower() != ".docx":
        parser.error("示例输出必须使用 .docx 扩展名。")

    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal.paragraph_format.line_spacing = 1.5
    doc.add_heading("MathType 公式转换示例", level=1)
    doc.add_paragraph("下面包含两条带编号的独立公式和一条行内公式。")
    doc.add_paragraph("[[MT:eq24]]")
    doc.add_paragraph("若 [[MT:inline1]]，则记录该项。")
    doc.add_paragraph("[[MT:eq1-1]]")

    output.parent.mkdir(parents=True, exist_ok=True)
    try:
        with output.open("xb") as stream:
            doc.save(stream)
    except FileExistsError:
        parser.error(f"已有示例文件，请指定新的输出路径：{output}")
    print(output)


if __name__ == "__main__":
    main()
