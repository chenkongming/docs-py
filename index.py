import docx
from docx import Document
import pandas as pd

sheet = pd.read_excel("name.xlsx")

def replaceWord(index,name):
    document=Document('temp.docx')
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            if "i" in run.text:
                run.text=run.text.replace('i', str(index))
            if "name" in run.text:
                run.text=run.text.replace('name',name)
    document.save("ouput/"+name+'.docx')

for row in sheet.values:
    replaceWord(row[0],row[1])





