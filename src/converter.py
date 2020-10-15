from docx2pdf import convert
import win32com.client
import sys
import os

# convert a single file
def c_file_DOCX_PDF(in_file_path, op_file_path = None):

    if op_file_path == None:
        convert(in_file_path)
    else:
        convert(in_file_path, op_file_path)

# convert all the files present in a folder
def c_batch_DOCX_PDF(in_folder_path):
    convert(in_folder_path)

#-----------------------------------------------------PPT to PDF-------------------------------------



# THIS IS NOT WORKING NEEDS TO BE FIXED


# Here formatType is an ENUM -> https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
def c_file_PPT_PDF(in_file_path, op_file_path, formatType = 32):

    # Convert file paths to Windows format
    in_file_path = os.path.abspath(in_file_path)
    op_file_path = os.path.abspath(op_file_path)

    print(in_file_path)
    print(op_file_path)

    powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
    powerpoint.Visible = 1

    slides = powerpoint.Presentations.Open(in_file_path)
    op_file_path = op_file_path + '\\output.pdf'
    print(op_file_path)
    slides.SaveAs(op_file_path, 32)
    slides.Close()
    powerpoint.Quit()