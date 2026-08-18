from pathlib import Path

from build_academy_ch1_docx import build


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "正文" / "正文MD" / "学院魔女_第五章_我愿折了旧杖，走向秋天.md"
OUTPUT = ROOT / "正文" / "学院魔女_第五章_我愿折了旧杖，走向秋天.docx"
TITLE = "第五章 我愿折了旧杖，走向秋天"


if __name__ == "__main__":
    build(SOURCE, OUTPUT, TITLE)
