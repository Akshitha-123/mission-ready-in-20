from PyPDF2 import PdfReader

reader = PdfReader("DD-Form-2977.pdf")
page = reader.pages[0]
print("MediaBox:", page.mediabox)
