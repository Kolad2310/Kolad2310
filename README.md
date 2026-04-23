```
import re
from docx import Document
from docx.shared import RGBColor

def write_2word_colored(paragraph, text):
    """
    Writes text into a Word paragraph with:
    - Green for positive values
    - Red for negative values
    - Applies to both $ values and bracket content
    """

    # Token pattern:
    # 1. Brackets → (....)
    # 2. Dollar values → $18.8m, $-20m
    # 3. Everything else
    tokens = re.findall(r'\(.*?\)|\$\-?\d+\.?\d*[a-zA-Z%]*|[^$()]+', text)

    for token in tokens:
        run = paragraph.add_run(token)

        # ----------------------
        # Bracket handling
        # ----------------------
        if token.startswith('(') and token.endswith(')'):
            if token.startswith('($'):
                run.font.color.rgb = RGBColor(255, 0, 0)  # Red
            else:
                run.font.color.rgb = RGBColor(0, 128, 0)  # Green

        # ----------------------
        # Dollar value handling
        # ----------------------
        elif token.startswith('$'):
            if '$-' in token:
                run.font.color.rgb = RGBColor(255, 0, 0)  # Red
            else:
                run.font.color.rgb = RGBColor(0, 128, 0)  # Green

        # बाकी normal text → default (no color)


def create_word_doc(text_list, output_file="output.docx"):
    """
    text_list: list of commentary strings
    """
    doc = Document()
    doc.add_heading("Commentary", 0)

    for text in text_list:
        p = doc.add_paragraph()
        write_2word_colored(p, text)

    doc.save(output_file)


