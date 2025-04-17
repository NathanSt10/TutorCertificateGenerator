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

    path = os.path.join(directory, name)
    certificate.save(path + ".pptx")


def main():
    path = os.getcwd()
    xlsx_path = os.path.join(path, "attendance")
    xlsx_files = [f for f in os.listdir(xlsx_path) if f.endswith('.xlsx')]

    if len(xlsx_files) == 1:
        file_path = os.path.join(xlsx_path, xlsx_files[0])
        workbook = openpyxl.load_workbook(file_path)
        print(f"Opened: {file_path}")
    else:
        print("Error: Expected exactly one .xlsx file in the directory.")
        exit(1)

    for sheet in workbook:
        if sheet.title == "Finished Level 3":
            continue
        os.makedirs(sheet.title, exist_ok=True)
        directory = os.path.join(path, sheet.title)
        template_name = os.path.join(path, "templates", sheet.title + ".pptx")
        certificate = Presentation(template_name)
        for row in sheet.iter_rows(min_row=2, max_col=1):
            for cell in row:
                if cell.value is not None and cell.font.b:
                    tutor_name = parse_name(cell.value)
                    create_certificate(certificate, tutor_name, directory)


if __name__ == "__main__":
    main()
