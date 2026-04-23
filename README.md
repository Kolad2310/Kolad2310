```
import re
from docx import Document
from docx.shared import RGBColor

def write_text_with_colors(text, output_file="output.docx"):
    doc = Document()
    p = doc.add_paragraph()

    # Pattern:
    # 1. Brackets → (....)
    # 2. Dollar values → $18.8m, $-20m, etc.
    pattern = r'(\(.*?\)|\$\-?\d+\.?\d*[a-zA-Z%]*)'

    parts = re.split(pattern, text)

    for part in parts:
        if not part:
            continue

        run = p.add_run(part)

        # ----------------------
        # Case 1: Bracket values
        # ----------------------
        if part.startswith('(') and part.endswith(')'):
            if part.startswith('($'):
                run.font.color.rgb = RGBColor(255, 0, 0)  # Red
            else:
                run.font.color.rgb = RGBColor(0, 128, 0)  # Green

        # ----------------------
        # Case 2: Dollar values
        # ----------------------
        elif part.startswith('$'):
            if '$-' in part:
                run.font.color.rgb = RGBColor(255, 0, 0)  # Red
            else:
                run.font.color.rgb = RGBColor(0, 128, 0)  # Green

        # ----------------------
        # Normal text → no color
        # ----------------------
        else:
            pass

    doc.save(output_file)

text = "Revenue was $18.8m (20.8% YoY), ($20m), ($30m/42%) vs last year."

write_text_with_colors(text)

