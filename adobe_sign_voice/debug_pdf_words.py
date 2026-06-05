"""Diagnostic: pdfminer text-box extraction + img2table 300-DPI PNG."""
from pathlib import Path
from pdfminer.high_level import extract_pages
from pdfminer.layout import LAParams, LTTextBox

pdf_path = r'C:\Users\Public\RPA\code\PSTN Migration\adobe_sign_voice\output\David Murphy_file_attachment_20260530_212341.pdf'

print("=" * 70)
print("STRATEGY E: pdfminer LTTextBox with tight char_margin=0.1")
print("=" * 70)

laparams = LAParams(
    char_margin=0.1,
    word_margin=0.05,
    line_margin=0.2,
    boxes_flow=None,
    detect_vertical=False,
)

boxes = []
for page_layout in extract_pages(pdf_path, laparams=laparams):
    page_h = page_layout.height
    for elem in page_layout:
        if isinstance(elem, LTTextBox):
            text = elem.get_text().strip()
            if not text:
                continue
            y_top = page_h - elem.bbox[3]
            boxes.append({
                'x0': round(elem.bbox[0], 1),
                'y_top': round(y_top, 1),
                'x1': round(elem.bbox[2], 1),
                'text': text.replace('\n', ' | '),
            })

boxes.sort(key=lambda b: (b['y_top'], b['x0']))
for b in boxes[:30]:
    print(f"  x0={b['x0']:6.1f}  x1={b['x1']:6.1f}  y={b['y_top']:6.1f}  {b['text'][:70]}")

print()
print("=" * 70)
print("STRATEGY F: img2table on saved 300-DPI PNG")
print("=" * 70)
png_path = r'C:\Users\Public\RPA\code\PSTN Migration\adobe_sign_voice\debug_page1_300dpi.png'
if Path(png_path).exists():
    try:
        from img2table.document import Image as Img2Image
        from img2table.ocr import EasyOCR
        ocr = EasyOCR(lang=["en"])
        for borderless in (False, True):
            img2 = Img2Image(src=png_path)
            table_map = img2.extract_tables(ocr=ocr, implicit_rows=True, borderless_tables=borderless, min_confidence=40)
            total = sum(len(t) for t in table_map.values())
            print(f"borderless={borderless}: tables found = {total}")
            for tables in table_map.values():
                for tbl in tables:
                    if tbl.df is not None and not tbl.df.empty:
                        print(f"  shape={tbl.df.shape}")
                        print(tbl.df.head(3).to_string())
    except Exception as e:
        import traceback; traceback.print_exc()
else:
    print("PNG not found - run previous debug first")
