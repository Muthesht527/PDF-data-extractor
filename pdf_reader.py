import PyPDF2
from pathlib import Path
from openai import images
import pytesseract as tess
from pytesseract import Output
from PIL import Image
import cv2
from pandas import DataFrame,merge

excel_file = 'student_record.xlsx'

#----Gets lines between two keywords----
def get_lines_between(lines, start_keyword, end_keyword):
    target=''
    start_index = end_index = None
    for i, line in enumerate(lines):
        if start_keyword.lower() in line.lower():
            start_index = i
        elif end_keyword.lower() in line.lower() and start_index is not None:
            end_index = i
            break
    if start_index is not None and end_index is not None and end_index > start_index:
        target = ' '.join(lines[start_index+1:end_index])
    return target

#----returns words between two keywords----
def read_page_by_word(reader, page_index, start_keyword, end_keyword):
    target=''
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

        if images is None:
            raise RuntimeError(
                "This PyPDF2 version does not support page.images. "
                "Upgrade PyPDF2 (3.x) or use another library."
            )

        img=cv2.imread(f"extracted_images/page_{page_index+1:03d}_img_001.jpeg") # read the image file using OpenCV

        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)    #Important for scanned PDF
        gray = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)[1]

        text = tess.image_to_string(gray)

        text_paras=text.split("\n\n")   #Splitting whole text into paragraphs
        
        for i, s in enumerate(text_paras):
            if start_keyword.lower() in s.lower() and end_keyword.lower() in s.lower():  #Searching which para contains the start and end keyword
                text=s
                break

        start_index=text.lower().rfind(start_keyword.lower())
        end_index=text.lower().find(end_keyword.lower(),start_index)

        # print(start_index,end_index,text[start_index:end_index].strip())      
        target=text[start_index+len(start_keyword):end_index].replace("\n"," ").strip(', ')

        return target

#----Returns lines between given keywords of a line----
def read_page_by_line(reader, page_index, start_keyword,end_keyword):
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

        if images is None:
            raise RuntimeError(
                "This PyPDF2 version does not support page.images. "
                "Upgrade PyPDF2 (3.x) or use another library."
            )

        img=cv2.imread(f"extracted_images/page_{page_index+1:03d}_img_001.jpeg") # read the image file using OpenCV

        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)    #Important for scanned PDF
        gray = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)[1]

        text = tess.image_to_string(gray)  # extract text from the image using Tesseract OCR

        lines=[line.strip() for line in text.splitlines() if line.strip()]  # split text into lines and remove empty lines

            # for i,line in enumerate(lines):
            #     print(f"{i}: {line}")

            # print("\n\n\n\n")
            
        return get_lines_between(lines, start_keyword, end_keyword)

output_dir = Path("extracted_images")  # create a Path object for the output folder
output_dir.mkdir(parents=True, exist_ok=True)  # create the folder if it doesn't exist
tess.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"  # set the path to the Tesseract executable

with open("Sample project pdf.pdf", "rb") as file:
    reader = PyPDF2.PdfReader(file)  # parse the PDF structure from the file stream
    req_pages=[0,2,len(reader.pages)-1] # list of page indices to extract images from (0-based)

    title=read_page_by_line(reader,req_pages[0],"an autonomous","project report")
    names=read_page_by_line(reader,req_pages[0],"submitted by","in partial fulfillment")
    details=names.replace(" ", "").replace(')', '(').split('(')
    names=details[:-2:2]
    reg_no=details[1::2]
    supervisor=read_page_by_word(reader,req_pages[1],"supervisor","professor")
    d1=DataFrame({'Name':names,'Reg No.':reg_no})
    d2=DataFrame({'Project Title':[title],'Supervisor':[supervisor]})
    d1=d1.merge(d2,how='cross')
    
    print(d1)

    d1.to_excel(excel_file, index=False)

    # for i,para in enumerate(page2_paras):    #Paragraph 8 contains 'supervisor' name
    #     print(f"{i}: {para}")
    

    # if len(reader.pages) == 0:  # total number of pages in the PDF
    #     print("This PDF has no pages.")
    # else:
    #     total = 0
    #     for page_index in req_pages:
    #         page = reader.pages[page_index]  # get the requested page (index 0-based)
    #         images = getattr(page, "images", None)  # list of embedded images on the page

    #         if images is None:
    #             raise RuntimeError(
    #                 "This PyPDF2 version does not support page.images. "
    #                 "Upgrade PyPDF2 (3.x) or use another library."
    #             )

    #         img=cv2.imread(f"extracted_images/page_{page_index+1:03d}_img_001.jpeg") # read the image file using OpenCV

    #         gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)    #Important for scanned PDF
    #         gray = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)[1]

    #         text = tess.image_to_string(gray)  # extract text from the image using Tesseract OCR

    #         lines=[line.strip() for line in text.splitlines() if line.strip()]  # split text into lines and remove empty lines

    #         for i,line in enumerate(lines):
    #             print(f"{i}: {line}")

    #         print("\n\n\n\n")
            
    #         print(get_lines_before(lines, "project report", 1))

            # print(text)  # print the extracted text
            # print("\n\n\n\n")

"--- Below part is the image saving part (A)---"
        #     for img_index, image in enumerate(images, start=1):  # loop images with 1-based index
        #         ext = getattr(image, "extension", None) or "jpeg"  # file extension for the image
        #         out_name = f"page_{page_index+1:03d}_img_{img_index:03d}.{ext}"
        #         out_path = output_dir / out_name
        #         with open(out_path, "wb") as out_file:
        #             out_file.write(image.data)  # raw bytes of the image
        #         total += 1

        # print(f"Saved {total} images from '{file.name}' to {output_dir}")
"--- (A) ---"


# def get_lines_after(lines, keyword, n):
#     for i, line in enumerate(lines):
#         if keyword.lower() in line.lower():
#             return lines[i+1:i+1+n]
#     return []

# def get_lines_before(lines, keyword, n):
#     target=''
#     for i, line in enumerate(lines):
#         if keyword.lower() in line.lower():
#             # return lines[max(0, i-n):i][0]  # return up to n lines before the keyword line
#             target = ' '.join(lines[max(0, i-n):i])
#     return target