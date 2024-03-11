# PrependPDF

PrependPDF is a metadata collation and PDF manipulation tool for the Schlesinger Library, Harvard Radcliffe Institute. This tool is used to automate via Python script the merging, metadata extraction, and cover sheet generation for DIP PDFs of born-digital documents. PrependPDF is intended for use in the pre-deposit phase of the born-digital workflow following document export and PDF conversion and followed by PDF/A conversion and validation

## Install

Install PrependPDF using pip install. Using PrependPDF depends on a font file, a CSS stylesheet, and internet connection to pull an institutional logo from Imgur. GhostScript's binary folder must be added to PATH in Windows Environmental Variables. GhostScript's PS2PDF.def file must be replaced with the file in this repository. The EXE distribution of PrependPDF is created using the following script:
pyinstaller --hidden-import pdfkit --hidden-import airium --hidden-import ghostscript --onefile PrependPDF<span>.p</span>y

The following libraries are required:
1. csv
2. pdfkit
3. os
4. shutil
5. PyPDF2
6. airium
7. reportlab
8. pdfrw
9. calendar
10. ghostscript
11. fitz
12. pandas
13. img2pdf

## Usage

In its source form, PrependPDF is only intended for use by Schlesinger Library staff for Schlesinger collections. Interested users may adapt the source code for their own use, but any products may not use the names of Schlesinger Library, Harvard Radcliffe Institute, or Harvard University.

## Contact

Zachary Maiorana
zachary_maiorana@radcliffe.harvard.edu
