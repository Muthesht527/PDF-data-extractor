import PyPDF2
from anyio import Path
import pandas as pd
import pytesseract as tess
from pytesseract import Output
import cv2
import re

tess.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"  # set the path to the Tesseract executable

def extract_image(reader, page_index, output_dir):
    if len(reader.pages) == 0:  # total number of pages in the PDF
        print("This PDF has no pages.")
    else:
        page = reader.pages[page_index]  # get the requested page (index 0-based)
        images = getattr(page, "images", None)  # list of embedded images on the page

        #Image extraction / saving part
        for img_index, image in enumerate(images, start=1):  # loop images with 1-based index
            ext = getattr(image, "extension", None) or "jpeg"  # file extension for the image
            out_name = f"page_{page_index+1:03d}_img_{img_index:03d}.{ext}"
            out_path = output_dir / out_name
            with open(out_path, "wb") as out_file:
                out_file.write(image.data)  # raw bytes of the image

def extract_title(structured_lines, anchor, max_gap=80):
    report_y = next(
        line['y'] for line in structured_lines
        if anchor in line['text'].lower()
    )

    above = sorted(
        [l for l in structured_lines if l['y'] < report_y],
        key=lambda x: x['y'],
        reverse=True
    )

    title_block = []
    prev_y = None

    for line in above:
        if prev_y is None or prev_y - line['y'] <= max_gap:
            title_block.append(line)
            prev_y = line['y']
        else:
            break

    return " ".join(l['text'] for l in reversed(title_block))

def extract_student(structured_lines,start_anchor,stop_anchor,max_gap=120):
    start_y = next(
        line['y'] for line in structured_lines
        if start_anchor in line['text'].lower()
    )

    candidates = sorted(
        [l for l in structured_lines if l['y'] > start_y],
        key=lambda x: x['y']
    )

    students = []
    prev_y = None

    for line in candidates:
        text = line['text'].lower()

        if stop_anchor in text:
            break

        if prev_y is None or line['y'] - prev_y <= max_gap:
            students.append(line['text'])
            prev_y = line['y']
        else:
            break

    return students

def Page1(page_index):
    students_split = []
    data=tess.image_to_data(f"extracted_images/page_{page_index+1:03d}_img_001.jpeg", output_type=Output.DATAFRAME)
    data = data[data.conf > 40]       # remove low confidence
    data = data[data.text.notna()]    # remove empty text
    data = data[data.text != " "]

    img=cv2.imread(f'extracted_images/page_{page_index+1:03d}_img_001.jpeg')

    lines = {}
    structured_lines = []

    for _, row in data.iterrows():
        x, y, w, h = row['left'], row['top'], row['width'], row['height']
        cv2.rectangle(img, (x, y), (x+w, y+h), (0,255,0), 1)
        line_id = (row['block_num'], row['par_num'], row['line_num'])
        lines.setdefault(line_id, []).append(row)


    for line in lines.values():
        line = sorted(line, key=lambda r: r['left'])
        text = " ".join([r['text'] for r in line])
        y = min(r['top'] for r in line)

        structured_lines.append({
            "text": text,
            "y": y
        })

    student_pattern = re.compile(r"(.+?)\s*\((\d+)\)")

    project_title = extract_title(structured_lines, "project report")
    students = extract_student(structured_lines, "submitted by", "in partial fulfillment")

    for s in students:
        m = student_pattern.search(s)
        if m:
            students_split.append({
                "name": m.group(1).strip(),
                "reg_no": m.group(2),
                "project_title": project_title
            })
    
    return students_split

def Page2(page_index):
    img = cv2.imread(f'extracted_images/page_{page_index+1:03d}_img_001.jpeg')

    if img is None:
        raise FileNotFoundError("Supervisor page image not found")

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)[1]

    data = tess.image_to_data(gray, output_type=Output.DATAFRAME)

    data = data[
        (data.conf > 40) &
        (data.text.notna()) &
        (data.text != "")
    ]

    lines = {}
    for _, row in data.iterrows():
        key = (row.block_num, row.par_num, row.line_num)
        lines.setdefault(key, []).append(row)

    structured_lines = []
    for line in lines.values():
        line = sorted(line, key=lambda r: r.left)
        text = " ".join(r.text for r in line)
        structured_lines.append(text)

    supervisor_idx = None

    for i, text in enumerate(structured_lines):
        if "supervisor" in text.lower():
            supervisor_idx = i
            break
    
    if supervisor_idx is not None:
        combined_text = structured_lines[supervisor_idx]

        # merge next line if it exists
        if supervisor_idx + 1 < len(structured_lines):
            combined_text += " " + structured_lines[supervisor_idx + 1]
    else:
        combined_text = None

    match = re.search(
        r"supervisor,\s*(dr\.\s*[a-z\.\s]+)",
        combined_text,
        re.IGNORECASE
    )

    if match:
        return match.group(1).strip()

    return None


with open("Sample project pdf.pdf", "rb") as file:
    reader = PyPDF2.PdfReader(file)  # parse the PDF structure from the file stream
    req_pages=[0,2,len(reader.pages)-1] # list of page indices to extract images from (0-based)
    output_dir = Path("extracted_images")  # create a Path object for the output folder
    
    page1=Page1(req_pages[0])
    print("Students:", page1)
    page2=Page2(req_pages[1])
    print("Supervisor:", page2)

# data=tess.image_to_data(f"extracted_images/page_{page_index+1:03d}_img_001.jpeg", output_type=Output.DATAFRAME)
# data = data[data.conf > 40]       # remove low confidence
# data = data[data.text.notna()]    # remove empty text
# data = data[data.text != " "]

# img=cv2.imread(f'extracted_images/page_{page_index+1:03d}_img_001.jpeg')

# lines = {}
# structured_lines = []

# for _, row in data.iterrows():
#     x, y, w, h = row['left'], row['top'], row['width'], row['height']
#     cv2.rectangle(img, (x, y), (x+w, y+h), (0,255,0), 1)
#     line_id = (row['block_num'], row['par_num'], row['line_num'])
#     lines.setdefault(line_id, []).append(row)


# for line in lines.values():
#     line = sorted(line, key=lambda r: r['left'])
#     text = " ".join([r['text'] for r in line])
#     y = min(r['top'] for r in line)

#     structured_lines.append({
#         "text": text,
#         "y": y
#     })

# student_pattern = re.compile(r"(.+?)\s*\((\d+)\)")
# students_split = []

# project_title = extract_title(structured_lines, "project report")
# students = extract_student(structured_lines, "submitted by", "in partial fulfillment")

# for s in students:
#     m = student_pattern.search(s)
#     if m:
#         students_split.append({
#             "name": m.group(1).strip(),
#             "reg_no": m.group(2),
#             "project_title": project_title
#         })

# print("Students:", students_split)
