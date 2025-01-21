from docxtpl import DocxTemplate
from docx2pdf import convert
import fitz  # PyMuPDF
import os


inch = "7".replace(",", ".")
# inch = "13".replace(",", ".")
resolution = "800x600"
# resolution = "1920x1080"
brightness = "600"
# touch = "резистивный"
touch = "проекционно-емкостной"

path = os.getcwd()
context = dict()

resolution_split = list(map(int, resolution.split("x")))

aspect_ratio = 1.34

if resolution_split[0] / resolution_split[1] > aspect_ratio:
    if "." in inch:
        template_file = "w_with_dot"
        inch_split = inch.split(".")
        context["inch_l"] = inch_split[0]
        context["inch_r"] = inch_split[1]
    else:
        template_file = "w"
        context["inch"] = inch
else:
    if "." in inch:
        template_file = "s_with_dot"
        inch_split = inch.split(".")
        context["inch_l"] = inch_split[0]
        context["inch_r"] = inch_split[1]
    else:
        template_file = "s"
        context["inch"] = inch

doc = DocxTemplate(f"{path}/templates/{template_file}.docx")

if "проекционно-емкостной" in touch:
    touch_value = "PCAP touch"
    filename_touch = "PCAP"
elif "резистивный" in touch:
    touch_value = "RES touch"
    filename_touch = "RES"

context.update(
    resolution_w=resolution_split[0],
    resolution_h=resolution_split[1],
    brightness=brightness,
    touch=touch_value,
)

doc.render(context)

output_doc_filename = f"{path}/output/{inch}-{context['resolution_w']}x{context['resolution_h']}-{context['brightness']}-{filename_touch}.docx"

doc.save(output_doc_filename)

output_pdf_filename = f"{path}/output/{inch}-{context['resolution_w']}x{context['resolution_h']}-{context['brightness']}-{filename_touch}.pdf"
convert(output_doc_filename, output_pdf_filename)

pdf_file = fitz.open(output_pdf_filename)
page = pdf_file.load_page(0)
pix = page.get_pixmap(dpi=600)
output_jpg_filename = f"{path}/output/{inch}-{context['resolution_w']}x{context['resolution_h']}-{context['brightness']}-{filename_touch}.jpg"
pix.save(output_jpg_filename)
pdf_file.close()

# os.remove(output_doc_filename)
# os.remove(output_pdf_filename)
