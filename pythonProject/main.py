from glob import glob
import re
import os
import win32com.client as win32
from win32com.client import constants

# Create list of paths to .doc files
paths = glob('C:\\Users\\anderson.peruci\\Desktop\\filesWordConverter\\*.doc', recursive=True)
# Create list of paths to .xls files
pathsXLS = glob('C:\\Users\\anderson.peruci\\Desktop\\filesWordConverter\\*.xls', recursive=True)

def save_as_docx(path):
    # Opening MS Word
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(path)
    doc.Activate ()

    # Rename path with .docx
    new_file_abs = os.path.abspath(path)
    new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)

    # Save and Close
    word.ActiveDocument.SaveAs(
        new_file_abs, FileFormat=constants.wdFormatXMLDocument
    )
    doc.Close(False)

def save_as_xlsx(path):
    # Opening MS Excel
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(path)

    # Save file as .xlsx
    wb.SaveAs(path + "x", FileFormat=51)

    # Close file and application
    wb.Close()
    excel.Application.Quit()

print("Conversion from .doc to .docx started...")
for path in paths:
    save_as_docx(path)

print("Conversion from .doc to .docx complete")
print("Conversion to .xls to .xlsx started...")

for path in pathsXLS:
    save_as_xlsx(path)

print("Conversion from .xls to .xlsx complete")
print("Completed conversions")