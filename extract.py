import os
import cv2
from matplotlib import pyplot as plt
from PyPDF2 import PdfReader
from pdf2image import convert_from_path
import pytesseract
from pytesseract import image_to_string
from PIL import Image, ImageFilter, ImageEnhance, ImageOps
import re
import numpy as np
import pandas as pd
import openpyxl
os.environ['TESSDATA_PREFIX'] = r'C:\Program Files\Tesseract-OCR\tessdata'


# Function to convert pdf to image
def convert_pdf_to_img(pdf_path):
  pdf_reader = PdfReader(pdf_path)
  total_pages = len(pdf_reader.pages)
  return convert_from_path(pdf_path, poppler_path = "C:\\Users\\User\\Desktop\\extract_voterlist\\poppler-23.08.0\\Library\\bin", dpi = 300, first_page=3, last_page=total_pages-2)


def image_processing(img):
  # Convert image to grayscale(black&white)
  img_gray = img.convert("L")
  # Contrast Increarse
  enhancer = ImageEnhance.Contrast(img_gray)
  factor = 2
  img0 = enhancer.enhance(factor)
  # Sharpness Increase
  enhancer = ImageEnhance.Sharpness(img0)
  factor = 2
  img1 = enhancer.enhance(factor)
  #image brightness enhancer
  enhancer = ImageEnhance.Brightness(img1)
  factor = 1.1
  img2 = enhancer.enhance(factor)
  return img2


# Function to crop the images
def crop_images(img, img_vill_cor, img_cor_rows):
  vill_name_img = img.crop(img_vill_cor)
  for row_index, row in enumerate(img_cor_rows):
    print("row:",row_index+1)
    for col_index, element in enumerate(row):
      print("column:",col_index)
      try:
        data_img = img.crop(element)
        texts = images_to_text(vill_name_img ,data_img)
        extract_data(texts)
      except ValueError as e:
        print(str(e))
  return data_dict


# Function to extract the data from the texts and store them in a dictionary
def extract_data(texts):
  lines = texts.split('\n')
  lines = [line for line in lines if line.strip() != '']
  print("Number of lines:", len(lines))  # Add this line for debugging
  # for i in range(len(lines)):
  #   if len(lines[i])==0:
  #     lines.pop(i)
  print("Lines:", lines)

  try:
    data = lines[0].split()
    # get section name
    sec_l = data[4:]
    # Join the elements into a string
    sec = ' '.join(sec_l)
    data_dict['Sec_no_vill'].append(sec)
  except IndexError:
    print("Index Error! 'None' value entered in the section key")
    data_dict['Sec_no_vill'].append(None)


  try:
     # get 2nd line
    data = lines[1].split()
    # get voter name
    name_l = data[3:]
    name = ' '.join(name_l)
    data_dict['Vtr_name'].append(name)
  except IndexError:
     print("Index Error! 'None' value entered in the voter_name key")
     data_dict['Vtr_name'].append(None)

  try:
    # get 3rd line
    data = lines[2].split()
    # get father/husband/others name
    ownr_n_l = data[3:]
    # Join the elements into a string
    ownr_n = ' '.join(ownr_n_l)
    data_dict['House_chief'].append(ownr_n)    
  except IndexError:
     print("Index Error! 'None' value entered in the House_chief key")
     data_dict['House_chief'].append(None)

  try:
    # get 4th line
    data = lines[3].split()
    # get house no
    hno_l = data[2:]
    hno = ' '.join(hno_l)
    data_dict['House_no'].append(hno)
  except IndexError:
     print("Index Error! 'None' value entered in the House_no key")
     data_dict['House_no'].append(None)

  try:
    # get 6th line
    data = lines[4].split()
    # get Age of voter
    try:
        result = int(data[1])
    except ValueError:
        try:
            result = int(data[2])
        except ValueError:
            result = None
    age = result
    data_dict['Age'].append(age)
  except IndexError:
     print("Index Error! 'None' value entered in the Age key")
     data_dict['Age'].append(None)

  try:
    # get Sex of voter
    sex = data[-1]
    data_dict['Sex'].append(sex)
  except IndexError:
     print("Index Error! 'None' value entered in the Sex key")
     data_dict['Sex'].append(None)

# Function to convert images to text
texts = ""
def images_to_text(vill_name_img ,data_img):
  text1 = pytesseract.image_to_string(vill_name_img, config='--psm 6', lang='Devanagari')
  text2 = pytesseract.image_to_string(data_img, config='--psm 6', lang='Devanagari')
  if text2.strip() == "" or text2 == "\x0c":
    raise ValueError("Image was blank / no text found!")
  texts = text1 +" "+ text2
  return texts


# list of tuples to be sent to the crop function
img_vill_cor = (45, 63, 773, 111)
img_cor_r1 = [(67, 185, 475, 345),(862, 185, 1300, 345),(1650, 185, 2100, 345)]
img_cor_r2 = [(67, 515, 475, 675),(862, 515, 1300, 675),(1650, 515, 2100, 675)]
img_cor_r3 = [(67, 845, 475, 1005),(862, 845, 1300, 1005),(1650, 845, 2100, 1005)]
img_cor_r4 = [(67, 1176, 475, 1336),(862, 1176, 1300, 1336),(1650, 1176, 2100, 1336)]
img_cor_r5 = [(67, 1506, 475, 1666),(862, 1506, 1300, 1666),(1650, 1506, 2100, 1666)]
img_cor_r6 = [(67, 1836, 475, 1996),(862, 1836, 1300, 1996),(1650, 1836, 2100, 1996)]
img_cor_r7 = [(67, 2169, 475, 2329),(862, 2169, 1300, 2329),(1650, 2169, 2100, 2329)]
img_cor_r8 = [(67, 2500, 475, 2660),(862, 2500, 1300, 2660),(1650, 2500, 2100, 2660)]
img_cor_r9 = [(67, 2830, 475, 2990),(862, 2830, 1300, 2990),(1650, 2830, 2100, 2990)]
img_cor_r10 = [(67, 3160, 475, 3320),(862, 3160, 1300, 3320),(1650, 3160, 2100, 3320)]
img_cor_rows = [img_cor_r1,img_cor_r2,img_cor_r3,img_cor_r4,img_cor_r5,img_cor_r6,img_cor_r7,img_cor_r8,img_cor_r9,img_cor_r10]


src_dir = 'D:/Amber_AC_Final_Revision/Amber_Final_PDFs'

dst_dir = 'D:/Amber_AC_Final_Revision/Amber_Final_Voter_List_Excel'

# Get PDF files from the source directory
pdf_files = [f for f in os.listdir(src_dir) if f.lower().endswith('.pdf')]

# Loop through PDF files and process
for pdf_file in pdf_files:
    print("File No:{}".format(pdf_file))
    pdf_path = os.path.join(src_dir, pdf_file)
    all_imgs = convert_pdf_to_img(pdf_path)
    data_dict = {
    "Sec_no_vill":[],
    "Vtr_name":[],
    "House_chief":[],
    "House_no":[],
    "Age":[],
    "Sex":[]
    }
    for i in range(len(all_imgs)):
        processed_image = image_processing(all_imgs[i])
        data = crop_images(processed_image, img_vill_cor, img_cor_rows)
        print("Page_{} Completed".format(i))

    print("Entire {} file data extracted!".format(pdf_file))

    # Create the Excel file path in the destination directory
    excel_file_name = pdf_file.replace('.pdf', '.xlsx')
    excel_path = os.path.join(dst_dir, excel_file_name)
    
    # Convert processed_data to a DataFrame (adjust this part based on your data structure)
    df = pd.DataFrame.from_dict(data_dict)
    
    # Save the DataFrame to Excel
    df.to_excel(excel_path, index=False)
    print("Entire {} file saved as {} file!".format(pdf_file, excel_file_name))
