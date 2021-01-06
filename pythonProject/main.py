import sys
from glob import glob
import re
import win32com.client as win32
import os
from win32com.client import constants
import shutil


# Create list of paths to .doc files
paths = glob('C:\\Users\\anderson.peruci\\Desktop\\filesWordConverter\\*.doc', recursive=True)
# Create list of paths to .xls files
pathsXLS = glob('C:\\Users\\anderson.peruci\\Desktop\\filesWordConverter\\*.xls', recursive=True)
i = 0

def save_as_docx(path):
    # Opening MS Word
    try:
        try:
            word = win32.gencache.EnsureDispatch('Word.Application')
        except AttributeError:
            # Remove cache and try again.
            MODULE_LIST = [m.__name__ for m in sys.modules.values()]
            for module in MODULE_LIST:
                if re.match(r'win32com\.gen_py\..+', module):
                    del sys.modules[module]
            shutil.rmtree(os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py'))
            from win32com import client
            word = win32.gencache.EnsureDispatch('Word.Application')

        # Open file without dialog box
        doc = word.Documents.OpenNoRepairDialog(path, False, True)
        # doc = word.Documents.Open(path)
        # doc.Activate()

        # Rename path with .docx
        new_file_abs = os.path.abspath(path)
        new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)

        # Save and Close
        word.ActiveDocument.SaveAs(
            new_file_abs, FileFormat=constants.wdFormatXMLDocument
        )
        doc.Close(False)

        # Move old file to another folder
        filename = os.path.basename(path)
        pathsold = "C:\\Users\\anderson.peruci\\Desktop\\Corrompido\\oldFiles\\" + filename
        shutil.move(path, pathsold)

    except:
        print("Unexpected error:", sys.exc_info()[0])

def save_as_docx_test(path):
    # Opening MS Word
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(path)
    doc.Activate()

    print("Word: " + str(word))
    print("Doc: " + str(doc))

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

    # Move old file to another path
    filename = os.path.basename(path)
    pathsold = "C:\\Users\\anderson.peruci\\Desktop\\Corrompido\\oldFiles\\" + filename
    shutil.move(path, pathsold)


print("Conversion from .doc to .docx started...")
for path in paths:
    save_as_docx(path)
    i += 1
    print(str(i) + " - " + str(path))

print("Conversion from .doc to .docx complete")
print("Converted to .docx: " + str(i))
print("Conversion to .xls to .xlsx started...")
i = 0

for path in pathsXLS:
    save_as_xlsx(path)
    i += 1
    print(str(i) + " - " + str(path))

print("Conversion from .xls to .xlsx complete")
print("Converted to .xlsx: " + str(i))
print("Completed conversions")