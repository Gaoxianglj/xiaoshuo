from pathlib import Path

from build_academy_ch1_docx import build


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "正文" / "正文MD" / "学院魔女_第四章_若我能为你照亮这黑夜.md"
OUTPUT = ROOT / "正文" / "学院魔女_第四章_若我能为你照亮这黑夜.docx"
TITLE = "第四章 若我能为你照亮这黑夜"


if __name__ == "__main__":
    build(SOURCE, OUTPUT, TITLE)
