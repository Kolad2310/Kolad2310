```
import re
from docx import Document
from docx.shared import RGBColor

def write_2word_colored(paragraph, text):
    """
    Writes text preserving exact spacing, while coloring:
    - ($...) → Red
    - (...)  → Green
    - $values → Green/Red based on sign
    """

    # Pattern for:
    # 1. Brackets
    # 2. Dollar values
    pattern = r'\(.*?\)|\$\-?\d+\.?\d*[a-zA-Z%/]*'

    last_idx = 0

    for match in re.finditer(pattern, text):
        start, end = match.span()

        # Add normal text BEFORE match (preserve spaces exactly)
        if start > last_idx:
            paragraph.add_run(text[last_idx:start])

        token = match.group()
        run = paragraph.add_run(token)

        # ----------------------
        # Bracket logic
        # ----------------------
        if token.startswith('('):
            if token.startswith('($'):
                run.font.color.rgb = RGBColor(255, 0, 0)  # Red
            else:
                run.font.color.rgb = RGBColor(0, 128, 0)  # Green

        # ----------------------
        # Dollar logic
        # ----------------------
        elif token.startswith('$'):
            if '$-' in token:
                run.font.color.rgb = RGBColor(255, 0, 0)  # Red
            else:
                run.font.color.rgb = RGBColor(0, 128, 0)  # Green

        last_idx = end

    # Add remaining text AFTER last match
    if last_idx < len(text):
        paragraph.add_run(text[last_idx:])


def create_word_doc(text_list, output_file="output.docx"):
    doc = Document()
    doc.add_heading("Commentary", 0)

    for text in text_list:
        p = doc.add_paragraph()
        write_2word_colored(p, text)

    doc.save(output_file)
