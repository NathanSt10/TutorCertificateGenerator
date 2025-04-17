import openpyxl
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt
import os


# raw_name: Last, First
def parse_name(raw_name):
    parts = raw_name.split(',')
    if len(parts) == 2:
        last_name = parts[0].strip()
        first_name = parts[1].strip()
        return f"{first_name} {last_name}"
    return parts


def create_certificate(certificate, name, directory):
    slide = certificate.slides[0]
    insert_name_tf = slide.shapes[4].text_frame
    insert_name_tf.clear()

    insert_name_tf.paragraphs[0].alignment = PP_ALIGN.CENTER

    p = insert_name_tf.paragraphs[0]
    p.text = name

    run = p.runs[0]
    run.font.size = Pt(40)
    run.font.bold = True
    run.font.name = "Century Schoolbook"

    path = directory + "/" + name
    certificate.save(path + ".pptx")


def main():
    path = os.getcwd()
    workbook = openpyxl.load_workbook("Attendance.xlsx")

    for sheet in workbook:
        if sheet.title == "Finished Level 3": continue
        os.makedirs(sheet.title, exist_ok=True)
        directory = path + "/" + sheet.title
        template_name = path + "/templates/" + sheet.title + ".pptx"
        certificate = Presentation(template_name)
        for row in sheet.iter_rows(min_row=2, max_col=1):
            for cell in row:
                if cell.value is not None and cell.font.b:
                    tutor_name = parse_name(cell.value)
                    create_certificate(certificate, tutor_name, directory)


if __name__ == "__main__":
    main()
