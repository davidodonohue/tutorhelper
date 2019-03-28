### Credit to https://stackoverflow.com/questions/6011115/doc-to-pdf-using-python

import sys
import os
import win32com.client

path = os.path.dirname(os.path.abspath(__file__))
current_path = path

wdFormatPDF = 17

for dir in os.listdir(path):
    if os.path.isdir(dir):
        current_path = os.path.join(path, dir)
        for file in os.listdir(current_path):
            print(file)
            in_file = os.path.join(current_path, file)
            out_file = os.path.join(current_path, file.split(".")[0] + ".pdf")
            word = win32com.client.Dispatch('Word.Application')
            doc = word.Documents.Open(in_file)
            doc.SaveAs(out_file, FileFormat=wdFormatPDF)
doc.Close()
word.Quit()