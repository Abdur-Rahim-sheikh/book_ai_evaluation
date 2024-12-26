from docx import Document

from document_section_processor import DocumentSectionProcessor

# use below if you want to take the help of paid openai
# from document_section_processor_2 import DocumentSectionProcessor

dsp = DocumentSectionProcessor()
src = "demo_90191.docx"
doc = Document(src)

dsp.pipeline(doc)

doc.save("new_modification_" + src.split("/")[-1])
