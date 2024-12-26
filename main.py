from docx import Document

from document_section_processor import DocumentSectionProcessor


dsp = DocumentSectionProcessor()
src = "demo_90191.docx"
doc = Document(src)

dsp.pipeline(doc)

doc.save("modified_" + src.split("/")[-1])
