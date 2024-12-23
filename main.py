from fontTools.ttLib import TTFont

from docx import Document
from unicode_font_converter import UnicodeFontConverter









ufc = UnicodeFontConverter()


doc = Document('resources/Project eBook Automation/Ebook/90191.docx')

cnt = 0
for para in doc.paragraphs:
    for run in para.runs:
        # run.text = ufc.convert(run.text, run.font.name)
        print(f"{run.text=}, {run.font.name=}")
        
        print(run.text)
    cnt+=1
    if cnt == 10:
        break

