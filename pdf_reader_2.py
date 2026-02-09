import PyPDF2
from pathlib import Path
import pandas as pd
import pytesseract as tess
from pytesseract import Output
import cv2
import re
import shutil
import hashlib

Input_DIR = Path("PDF_Input")
Processed_DIR = Path("Processed_PDFs")
Temp_DIR = Path("Temp_Img")

Processed_DIR.mkdir(exist_ok=True)
Temp_DIR.mkdir(exist_ok=True)

pdfs=list(Input_DIR.glob("*.pdf"))

#tesseract must be installed to the PATH or in the program files
# tess.pytesseract.tesseract_cmd = shutil.which("tesseract") or r"C:\Program Files\Tesseract-OCR\tesseract.exe"  # set the path to the Tesseract executable
excel_file='student_record.xlsx'

def extract_images(reader, page_index, output_dir):
    if len(reader.pages) == 0:  # total number of pages in the PDF
        print("This PDF has no pages.")
    else:
        page = reader.pages[page_index]  # get the requested page (index 0-based)
        images = getattr(page, "images", None)  # list of embedded images on the page

        #Image extraction / saving part
        if not images:
            return []

        saved = []
        for img_index, image in enumerate(images, start=1):  # loop images with 1-based index
            ext = getattr(image, "extension", None) or "jpeg"  # file extension for the image
            out_name = f"page_{page_index+1:03d}_img_{img_index:03d}.{ext}"
            out_path = output_dir / out_name
            with open(out_path, "wb") as out_file:
                out_file.write(image.data)  # raw bytes of the image
            saved.append(out_path)

        return saved

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

def Page1(reader, page_index, output_dir):
    students_split = []

    if len(reader.pages) == 0:  # total number of pages in the PDF
        print("This PDF has no pages.")
    else:
        output_dir.mkdir(parents=True, exist_ok=True)
        saved = extract_images(reader, page_index, output_dir)
        if not saved:
            raise FileNotFoundError("No images found on page 1")

        img_path = saved[0]

        data = tess.image_to_data(str(img_path), output_type=Output.DATAFRAME)
        
        data["conf"] = pd.to_numeric(data["conf"], errors="coerce")
        data = data[data.conf > 40]       # remove low confidence
        data = data[data.text.notna()]    # remove empty text
        data = data[data.text.str.strip() != ""]

        img = cv2.imread(str(img_path))

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

        project_title = extract_title(structured_lines, "project report")
        students = extract_student(structured_lines, "submitted by", "in partial fulfillment")

        student_pattern = re.compile(r"(.+?)\s*\((\d+)\)")

        for s in students:
            m = student_pattern.search(s)
            if m:
                students_split.append({
                    "Name": m.group(1).strip(),
                    "Register Number": m.group(2),
                    "Project Title": project_title
                })
        
        return students_split

def Page2(reader, page_index, output_dir):

    if len(reader.pages) == 0:  # total number of pages in the PDF
        print("This PDF has no pages.")
    else:
        output_dir.mkdir(parents=True, exist_ok=True)
        saved = extract_images(reader, page_index, output_dir)
        if not saved:
            raise FileNotFoundError("No images found on supervisor page")

        img = cv2.imread(str(saved[0]))

        if img is None:
            raise FileNotFoundError("Supervisor page image not found")

        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        gray = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)[1]

        data = tess.image_to_data(gray, output_type=Output.DATAFRAME)

        data["conf"] = pd.to_numeric(data["conf"], errors="coerce")
        data = data[
            (data.conf > 40) &
            (data.text.notna()) &
            (data.text.str.strip() != "")
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

        if combined_text:
            match = re.search(
                r"supervisor,\s*(dr\.\s*[a-z\.\s]+)",
                combined_text,
                re.IGNORECASE
            )
        else:
            match = None

        if match:
            return match.group(1).strip()

        return None
  
def Cert_Page(reader, page_index, output_dir, student_names=None, college_keywords=None):
    if student_names is None:
        student_names = []

    if college_keywords is None:
        college_keywords = ["rajalakshmi", "anna university"]

    output_dir.mkdir(parents=True, exist_ok=True)
    saved = extract_images(reader, page_index, output_dir)
    if not saved:
        raise FileNotFoundError("No images found on certificate page")

    img = cv2.imread(str(saved[0]))
    if img is None:
        raise FileNotFoundError("Certificate page image not found")

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)[1]

    h, w = gray.shape

    data = tess.image_to_data(gray, output_type=Output.DATAFRAME)

    data["conf"] = pd.to_numeric(data["conf"], errors="coerce")
    data = data[
        (data.conf > 40) &
        (data.text.notna()) &
        (data.text.str.strip() != "")
    ]

    # Normalize text
    data["text"] = data["text"].str.lower()

    certificate_hits = data[data.text.str.contains("certificate", na=False)]

    text_match = len(certificate_hits) > 0

    layout_match = False
    large_text_match = False

    for _, row in certificate_hits.iterrows():
        # near top of page
        if row.top < 0.25 * h:
            layout_match = True

        # large text (heading)
        if row.height > 40:
            large_text_match = True

    confidence = sum([
        text_match,
        layout_match,
        large_text_match
    ])

    certificate_present = confidence >= 2

    # ---------------------------
    # 4. Optional validity checks
    # ---------------------------
    # full_text = " ".join(data.text.tolist())

    # validity = {
    #     "student_name": any(
    #         name.lower() in full_text
    #         for name in student_names
    #     ),
    #     "college": any(
    #         kw in full_text
    #         for kw in college_keywords
    #     ),
    #     "signature": any(
    #         w in full_text
    #         for w in ["principal", "hod", "signature", "chairman"]
    #     ),
    #     "date": bool(
    #         re.search(r"\b(20\d{2})\b", full_text)
    #     )
    # }

    return {
        "certificate_present": certificate_present,
        "confidence": confidence,
        "text_match": text_match,
        "layout_match": layout_match,
        # "validity_checks": validity
    }

def process_single_pdf(pdf_path):
    print(f"Processing: {pdf_path.name}")

    # --- Clean temp directory ---
    if Temp_DIR.exists():
        shutil.rmtree(Temp_DIR)
    Temp_DIR.mkdir()

    with open(pdf_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        req_pages = [0, 2, len(reader.pages) - 1]

        page1 = Page1(reader, req_pages[0], Temp_DIR)
        supervisor = Page2(reader, req_pages[1], Temp_DIR)
        cert_page = Cert_Page(reader,req_pages[2],Temp_DIR,student_names=[s['Name'] for s in page1])

        pdf_id=pdf_hash(pdf_path)

        page1 = pd.DataFrame(page1)
        page1["Register Number"] = (
            page1["Register Number"].astype(str).str.strip().str.replace(r"\D", "", regex=True)
        )
        page1["PDF_ID"] = pdf_id
        page1["PDF Name"] = pdf_path.name
        page2 = pd.DataFrame([{
            "Supervisor": supervisor,
            "Certificate Status": cert_page["certificate_present"],
        }])

        data = page1.merge(page2, how="cross")

    # --- Move processed PDF ---
    shutil.move(pdf_path,Processed_DIR / pdf_path.name)

    # --- Cleanup temp images ---
    shutil.rmtree(Temp_DIR)

    print(f"✔ Done: {pdf_path.name}")
    return data

def Upload_Excel(data, excel_file):
    if data:
        final_df = pd.concat(data, ignore_index=True)

        if Path(excel_file).exists():
            old_df = pd.read_excel(excel_file)
            final_df = pd.concat([old_df, final_df], ignore_index=True)

        # Prevent duplicate PDFs
        final_df.drop_duplicates(
            subset=["PDF_ID", "Register Number"],
            inplace=True
        )

        final_df.to_excel(excel_file, index=False)

    print("✅ ALL PDFs PROCESSED")

def pdf_hash(pdf_path):
    h = hashlib.sha256()
    with open(pdf_path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


all_data = []

for pdf_path in pdfs:
    try:
        df = process_single_pdf(pdf_path)
        all_data.append(df)
    except Exception as e:
        print(f"❌ Failed: {pdf_path.name}")
        print(e)

Upload_Excel(all_data, excel_file)


# with open("Sample project pdf.pdf", "rb") as file:
#     reader = PyPDF2.PdfReader(file)  # parse the PDF structure from the file stream
#     req_pages=[0,2,len(reader.pages)-1] # list of page indices to extract images from (0-based)
#     output_dir = Path("extracted_images")  # create a Path object for the output folder
    
#     page1 = Page1(reader, req_pages[0], output_dir)
#     # print("Students:", page1)
#     page2 = Page2(reader, req_pages[1], output_dir)
#     # print("Supervisor:", page2)
#     cert_page = Cert_Page(reader, req_pages[2], output_dir, student_names=[s['Name'] for s in page1])
#     # print("Certificate Page Analysis:", cert_page)

#     page2 = {'Supervisor': page2,'Certificate Status': cert_page['certificate_present']} if page2 else {}

#     page1=pd.DataFrame(page1)
#     page2=pd.DataFrame([page2]) if page2 else pd.DataFrame()

#     upload=page1.merge(page2, how='cross')
#     # print(upload)
#     print("Upload done!!!")
    
#     upload.to_excel(excel_file, index=False)

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
