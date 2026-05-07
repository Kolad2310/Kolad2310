```
def write_colored_commentary(paragraph, text):

    pattern = r'''
        \(-?\$?\d[\d,]*\.?\d*[%mMkKbB]?\) |
        -?\$?\d[\d,]*\.?\d*[%mMkKbB]?
    '''

    parts = re.split(
        f'({pattern})',
        text,
        flags=re.VERBOSE
    )

    for part in parts:

        if not part:
            continue

        run = paragraph.add_run(part)

        nums = re.findall(
            r'-?\d[\d,]*\.?\d*',
            part
        )

        if nums:

            value = float(
                nums[0].replace(',', '')
            )

            # NEGATIVE = RED
            if '-' in part:

                run.font.color.rgb = RGBColor(
                    255, 0, 0
                )

            # POSITIVE = GREEN
            else:

                run.font.color.rgb = RGBColor(
                    0, 128, 0
                )
