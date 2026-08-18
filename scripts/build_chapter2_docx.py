from pathlib import Path

from build_academy_ch1_docx import build


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "正文" / "正文MD" / "学院魔女_第二章_你予我花与晨光.md"
OUTPUT = ROOT / "正文" / "学院魔女_第二章_你予我花与晨光.docx"
TITLE = "第二章 你予我花与晨光"


if __name__ == "__main__":
    build(SOURCE, OUTPUT, TITLE)
