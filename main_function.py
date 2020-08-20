import os
from pathlib import Path

# Importing pdf_downloader function
import pdfDownload
# Importing pre_process_data function
import data_prepocessing
# Importing generate_excel function
import excel_writer_module


# Removing PDF files from the PDFs Folder
def remove_pdfs():
    # Getting absolute path of current directory
    mypath = Path().absolute()

    # Defining path of the files
    pdfpath = mypath / 'All_Pdf_Files'

    if pdfpath.exists():
        for currentfilename in os.listdir(pdfpath):
            try:
                os.remove(pdfpath / currentfilename)
            except Exception as e:
                print('File is in use')

    return "Pdf's removed successfully"


if __name__ == '__main__':
    # Downloading PDF's
    pdfDownload.pdf_downloader()
    # Pre-process the data
    # data_prepocessing.pre_process_data()
    # Remove downloaded pdf's
    # remove_pdfs()
    # Generate Excel report
    # excel_writer_module.generate_excel(filename='MHP FEE ESTIMATE Excel Report.xlsx', foldername='Excel_files')
