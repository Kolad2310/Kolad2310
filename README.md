```
import re
from docx import Document
from docx.shared import RGBColor

def add_colored_text(paragraph, text):

    # Improved pattern (captures FULL expression properly)
    pattern = r'\(\s*\$?-?\d+\.?\d*m?(?:\s*/\s*-?\d+\.?\d*%?)?(?:\s*%?\s*YoY)?\s*\)|\$-?\d+\.?\d*m?(?:\s*/\s*-?\d+\.?\d*%?)?'

    parts = re.split(f'({pattern})', text)

    for part in parts:
        if not part:
            continue

        run = paragraph.add_run(part)

        # Apply color only to matched financial expressions
        if re.fullmatch(pattern, part.strip()):
            if '-' in part:
                run.font.color.rgb = RGBColor(255, 0, 0)  # RED
            else:
                run.font.color.rgb = RGBColor(0, 128, 0)  # GREEN


def write_to_word(text, filename="output.docx"):
    doc = Document()
    p = doc.add_paragraph()

    add_colored_text(p, text)

    doc.save(filename)


# Your text
text = """Europe TRI up by $30.5m (3.3% YoY), growth was led by Sweden ($8.7m) driven by IWPB / Others ($8.9m/118.8%), 
but offset by France ($-2.7m) driven by TxB ($-5.1m/-4.7%)"""

write_to_word(text)
