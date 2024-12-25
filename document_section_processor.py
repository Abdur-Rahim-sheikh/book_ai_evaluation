from docx import Document
from docx.section import Section
from docx.text.paragraph import Paragraph
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_ORIENTATION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


import logging
from io import BytesIO
import matplotlib.pyplot as plt
import pandas as pd
import cv2
import ollama
import numpy as np

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class DocumentSectionProcessor:
    CUSTOM_MODEL = "rokomari_bot"

    def __init__(self, modelfile_location: str = "static/Modelfile"):
        if self.CUSTOM_MODEL not in ollama.list():
            try:
                modelfile = open(modelfile_location, "r").read()
                ollama.create(self.CUSTOM_MODEL, modelfile=modelfile)
                logger.info(f"Model {self.CUSTOM_MODEL} created successfully")
            except Exception as e:
                logger.error(e)
            

    def pipeline(self, doc: Document):
        # the calling order is important
        for section in doc.sections:
            self.adjust_page_layout(section)
            self.remove_headers_footers(section)

        for paragraph in doc.paragraphs:
            self.enforce_content_restrictions(paragraph)
        # applying on whole document
        self.process_tables(doc)
        # self.formate_image(doc)

    # Function to adjust page layout
    def adjust_page_layout(self, section: Section):

        cols = section._sectPr.xpath("./w:cols")[0]
        cols.set(qn("w:num"), "1")
        section.orientation = WD_ORIENTATION.PORTRAIT
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

        page_borders = section._sectPr.xpath("./w:pgBorders")
        if page_borders:
            for border in page_borders:
                section._sectPr.remove(border)

    # Function to process headers and footers
    def remove_headers_footers(self, section: Section):
        header = section.header
        footer = section.footer
        header.is_linked_to_previous = False
        footer.is_linked_to_previous = False
        for paragraph in header.paragraphs:
            paragraph.clear()
        for paragraph in footer.paragraphs:
            paragraph.clear()

    # Function to remove unwanted content
    def enforce_content_restrictions(self, paragraph: Paragraph):
        text = paragraph.text.strip()
        if not text:
            return
        refined_text = ollama.generate(model=self.CUSTOM_MODEL, prompt=text)

    # Function to process fonts
    def fix_fonts(doc):
        for para in doc.paragraphs:
            for run in para.runs:
                if "Arabic" in run.text:
                    run.font.name = "Al Majeed Quranic"
                elif "Bengali" in run.text:
                    run.font.name = "SutonnyMJ"
                else:
                    run.font.name = "Times New Roman"
                run.font.size = Pt(16)

    def process_tables(self, doc: Document, max_width: Inches = Inches(3)):
        # process this at last as this will work on whole table
        for table in doc.tables:
            table_width = sum(cell.width for cell in table.rows[0].cells)
            logger.info(f"{table_width=}")
            if table_width <= max_width:
                continue

            table_data = [[cell.text for cell in row.cells] for row in table.rows]
            image_stream = self.array_to_image(table_data)

            parent = table._element.getparent()
            table_index = list(parent).index(table._element)
            parent.remove(table._element)
            new_paragraph_element = OxmlElement("w:p")
            parent.insert(
                table_index, new_paragraph_element
            )  # Insert at the same index
            paragraph = doc.paragraphs[table_index]

            # Add the image to the new paragraph
            run = paragraph.add_run()
            run.add_picture(image_stream)

    def array_to_image(self, table: list):
        df = pd.DataFrame(table[1:], columns=table[0])
        fig, ax = plt.subplots(figsize=(6, 2))
        ax.axis("tight")
        ax.axis("off")
        table = ax.table(
            cellText=df.values, colLabels=df.columns, loc="center", cellLoc="center"
        )
        image_buffer = BytesIO()
        fig.savefig(image_buffer, format="png", bbox_inches="tight")
        image_buffer.seek(0)
        return image_buffer

    def formate_image(self, doc: Document):
        for shape in doc.inline_shapes:
            rel_id = shape._inline.graphic.graphicData.pic.blipFill.blip.embed
            image_part = doc.part.related_parts[rel_id]

            original_image_bytes = image_part.blob
            new_image_bytes = self.__ushape_image(original_image_bytes)
            image_part.blob = new_image_bytes

    def __ushape_image(self, image_bytes: bytes):
        # dummy code
        np_arr = np.frombuffer(image_bytes, dtype=np.uint8)

        # 2) Decode the array into an image
        img = cv2.imdecode(np_arr, cv2.IMREAD_COLOR)  # BGR format

        # 3) Convert the image to grayscale
        gray_img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

        # 4) Encode the grayscale image back to bytes (PNG here)
        success, buf = cv2.imencode(".png", gray_img)

        # buf is a 1D array of bytes; we convert it to python bytes
        return buf.tobytes()

    # Function to clean text boxes
    def remove_text_boxes(doc):
        for shape in doc.inline_shapes:
            if shape.type == 3:  # Text box type
                # Extract text and re-add to the doc if necessary
                pass


# # Apply all processing functions
# adjust_page_layout(doc)
# remove_headers_footers(doc)
# enforce_content_restrictions(doc)
# fix_fonts(doc)
# process_tables(doc)
# process_images(doc)
# remove_text_boxes(doc)

# # Save the processed file
# doc.save('processed_book.docx')
