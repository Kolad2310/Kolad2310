```
import re
from docx import Document
from docx.shared import RGBColor

def add_colored_text(paragraph, text):
    # Pattern to capture numbers with $, %, brackets, / etc.
    pattern = r'\(?\$?-?\d+\.?\d*m?(?:/\-?\d+\.?\d*%?)?(?:\s*YoY)?\)?|\(?-?\d+\.?\d*%.*?\)?'

    parts = re.split(f'({pattern})', text)

    for part in parts:
        if not part:
            continue

        run = paragraph.add_run(part)

        # Apply color only to numeric patterns
        if re.fullmatch(pattern, part.strip()):
            if '-' in part:
                run.font.color.rgb = RGBColor(255, 0, 0)  # RED (negative)
            else:
                run.font.color.rgb = RGBColor(0, 128, 0)  # GREEN (positive)


def write_to_word(text, filename="output.docx"):
    doc = Document()
    p = doc.add_paragraph()

    add_colored_text(p, text)

    doc.save(filename)


# 🔹 Your input string
text = """Europe TRI up by $30.5m (3.3% YoY), growth led by Sweden ($8.7m) driven by IWPB / Others ($8.9m/118.8%), 
but offset by France ($-2.7m) driven by TxB ($-5.1m/-4.7%)"""

write_to_word(text)
