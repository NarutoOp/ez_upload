from docx2pdf import convert

# convert a single file
def convert_file(in_file_path, op_file_path = None):

    if op_file_path == None:
        convert(in_file_path)
    else:
        convert(in_file_path, op_file_path)

# convert all the files present in a folder
def convert_batch(in_folder_path):
    convert(in_folder_path)