import pandas as pd
import pytesseract as tess
from pytesseract import Output
import cv2
import re

# text='Muthesh (2117240070196) Satyarth (2117240070200)'
# print(text.replace(" ", "").replace(')', '(').split('('))

# d1=pd.DataFrame({'Name':['Muthesh','Satyarth'],'Reg No.':['2117240070196','2117240070200']})
# d2=pd.DataFrame({'Project Title':['An Autonomous Vehicle'],'Supervisor':['Dr. Smith']})

# d1=d1.merge(d2,how='cross')

# print(d1)


tess.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"  # set the path to the Tesseract executable
data=tess.image_to_data('extracted_images/page_001_img_001.jpeg', output_type=Output.DATAFRAME)
data = data[data.conf > 40]       # remove low confidence
data = data[data.text.notna()]    # remove empty text
data = data[data.text != " "]

img=cv2.imread('extracted_images/page_001_img_001.jpeg')

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
students_split = []

# cv2.imshow("OCR Boxes", img)
# cv2.waitKey(0)                #Too big to display
# cv2.destroyAllWindows()

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

# print(structured_lines)
# print("Project Title:", project_title)
print("Students:", students_split)


# cv2.imwrite('ocr_output.jpeg', img)

# print(data.head(30))