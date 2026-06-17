from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


def set_font(style, latin: str, east_asia: str, size: float, bold: bool = False, color: str | None = None) -> None:
    style.font.name = latin
    style._element.rPr.rFonts.set(qn("w:eastAsia"), east_asia)
    style.font.size = Pt(size)
    style.font.bold = bold
    if color:
        style.font.color.rgb = RGBColor.from_string(color)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("output")
    args = parser.parse_args()
    output = Path(args.output).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    section.top_margin = section.bottom_margin = Cm(2.54)
    section.left_margin = section.right_margin = Cm(2.54)
    normal = doc.styles["Normal"]
    set_font(normal, "Calibri", "等线", 11)
    normal.paragraph_format.line_spacing = 1.45
    normal.paragraph_format.space_after = Pt(6)
    for name, size, color in (("Title", 20, "1F4E79"), ("Heading 1", 15, "1F4E79"), ("Heading 2", 12, "365F91")):
        set_font(doc.styles[name], "Calibri", "等线", size, True, color)
    doc.styles["Title"].paragraph_format.space_after = Pt(16)
    doc.styles["Heading 1"].paragraph_format.space_before = Pt(14)
    doc.styles["Heading 1"].paragraph_format.space_after = Pt(6)
    doc.styles["Heading 1"].paragraph_format.keep_with_next = True
    doc.styles["Heading 2"].paragraph_format.space_before = Pt(10)
    doc.styles["Heading 2"].paragraph_format.space_after = Pt(4)
    doc.styles["Heading 2"].paragraph_format.keep_with_next = True
    doc.core_properties.title = "投资会议纪要模板"
    doc.core_properties.subject = "修正转录、Q&A与正式投资会议纪要"
    doc.save(output)


if __name__ == "__main__":
    main()
