from pathlib import Path

import build_stockton_ch3_docx as base
from docx_fangsong import embed_fangsong_docx


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "正文" / "正文MD" / "第四章_我愿忘记，将此当作你我初次相遇.md"
OUTPUT = ROOT / "正文" / "第四章_我愿忘记，将此当作你我初次相遇.docx"
def build():
    doc = base.Document()
    base.configure_document(doc)
    doc.core_properties.title = "第四章 我愿忘记，将此当作你我初次相遇"
    doc.core_properties.subject = "《学院魔女》斯托克顿篇正文"
    doc.core_properties.author = ""

    for line in SOURCE.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.startswith("# "):
            base.add_title(doc, stripped[2:].strip())
        elif stripped.startswith("> "):
            base.add_epigraph(doc, stripped[2:].strip())
        elif stripped.startswith("## "):
            base.add_section_heading(doc, stripped[3:].strip())
        elif stripped.startswith("*") and stripped.endswith("*"):
            base.add_body(doc, stripped[1:-1])
            doc.paragraphs[-1].runs[-1].italic = True
        else:
            base.add_body(doc, stripped)

    # 收紧结尾短段，避免最后一句被单独挤到空白尾页。
    for paragraph in doc.paragraphs[-12:]:
        paragraph.paragraph_format.space_after = base.Pt(4)
    for paragraph in doc.paragraphs[-4:]:
        paragraph.paragraph_format.space_after = base.Pt(0)

    doc.save(OUTPUT)
    embed_fangsong_docx(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    build()
