from PyPDF2 import PdfFileMerger
import glob
from pathlib import Path

def combine(path):
    pdfs = glob.glob(f"{path}\*.pdf")
    merger = PdfFileMerger()

    for pdf in pdfs:
        merger.append(pdf)

    merger.write(str(Path('.') / "out" / "result.pdf"))
    merger.close()

