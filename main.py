from docx import Document
from bangla_to_unicode import BanglaToUnicode
from document_section_processor import DocumentSectionProcessor

ufc = BanglaToUnicode()
dsp = DocumentSectionProcessor()
src = "demo_90191.docx"
doc = Document(src)

dsp.pipeline(doc)

doc.save("modified_" + src.split("/")[-1])
