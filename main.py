from docxtpl import DocxTemplate
from docx2pdf import convert
import fitz  # PyMuPDF
import os


# inch = input("Диагональ (например 13.3 или 13,3 или 7): ")
# resolution = input("Разрешение (например 800x480 или 800*480): ")
# brightness = input("Яркость (целое число): ")
# touch = input("Тип сенсорного экрана:\n1. PCAP\n2. 5W RES\n3. 4W RES\nВариант: ")

path = os.getcwd()

if "." in inch or "," in inch:
    



doc = DocxTemplate(f"{path}/w.docx")

context = {
    "inch": "7.4",
    "resolution_w": "1920",
    "resolution_h": "1200",
    "brightness": "400",
    "touch": "PCAP touch",
}

doc.render(context)

touch = ""

if "touch" in context:
    touch = "-PCAP"
    
output_doc_filename = f"{path}/output/{context['inch']}-{context['resolution_w']}x{context['resolution_h']}-{context['brightness']}{touch}.docx"

doc.save(output_doc_filename)

output_pdf_filename = f"{path}/output/{context['inch']}-{context['resolution_w']}x{context['resolution_h']}-{context['brightness']}{touch}.pdf"
convert(output_doc_filename, output_pdf_filename)

pdf_file = fitz.open(output_pdf_filename)
page = pdf_file.load_page(0)
pix = page.get_pixmap()
output_jpg_filename = f"{path}/output/{context['inch']}-{context['resolution_w']}x{context['resolution_h']}-{context['brightness']}{touch}.jpg"
pix.save(output_jpg_filename)
pdf_file.close()

# os.remove(output_doc_filename)
os.remove(output_pdf_filename)