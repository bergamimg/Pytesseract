import os
from PIL import Image
import pytesseract
from pathlib import Path
import re
import xlsxwriter
import pandas as pd

## Define Pytesseract.exe path file
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

## Define Image path file
path_file = r'C:\Desktop\GitHub\1.OCR'
list_documents = []
list_file_names = []
for file in os.listdir(path_file):
    if file.endswith(".png"): # or file.endswith(".JPG") :
        path_file_join = os.path.join(path_file, file)

        image_text = pytesseract.image_to_string(Image.open(path_file_join), lang='por')
        print("\n ", image_text, "\n")

        image_text = image_text.replace("\n", "")

        list_doc_per_file = re.findall(r"\d\d\d+[\.]+\d\d\d+[\.]+\d\d\d+[\-]+\d\d", image_text)
 
        strings = image_text
        final_results = [re.findall(r"(?<=\d)[A-Z](?=\d)", s) for s in strings]
        
        print("\n\n Final RegEx results: \n\n", final_results)
        lista_docs.append(list_doc_per_file)
        list_file_names.append(file)

union_lists = sum(list_documents, [])

print("\n\n lista de documentos appends: \n\n", list_file_names)

path_result_excel = r'C:\Desktop\GitHub\1.OCR\List_extracted_docs.xlsx'
dictionary_lists =  {'List_Extracted_Docs': union_lists}
dictionary_to_df = pd.DataFrame.from_dict(dictionary_lists)
dictionary_to_df.to_excel(path_result_excel, index = None, engine = 'xlsxwriter')
