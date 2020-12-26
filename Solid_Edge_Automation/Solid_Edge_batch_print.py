# the program creates a single pdf file from the .dft files
# avalaible in a directory filtering according to a list 

# for the conversion from .dft to .pdf it has been used the 
# SolidEdgeTranslationServices.exe service
# available at 
# C:\Program Files\Siemens\Solid Edge 2020\Program


import os
from PyPDF2 import PdfFileMerger

# the program takes as input the folder in which all project files
# and the .txt list for filtering are indicated 
fold = input("Specify project folder path   ")
lista = input("Specify list file name   ")

with open(fold + "\\" + lista + ".txt") as f:
    lines = f.readlines()

codici = [codice.replace('\n','') for codice in lines if codice!='\n']

comando = '"C:\Program Files\Siemens\Solid Edge 2020\Program\SolidEdgeTranslationServices.exe" -i='
opzioni = "-t=pdf " + "-m=TRUE"
# -i input
# -o output
# -t type
# -m multiple sheets
for codice in codici:
    os.system(comando + fold + "\\"+ codice + ".dft -o=" + fold + "\\" + codice + ".pdf " + opzioni)
    
merger = PdfFileMerger()

for codice in codici:
    pdf = fold + "\\" + codice + ".pdf"
    merger.append(open(pdf, 'rb'))

with open(fold + "\\" + "batch_print_" + lista + ".pdf", "wb") as fout:
    merger.write(fout)

merger.close()    
