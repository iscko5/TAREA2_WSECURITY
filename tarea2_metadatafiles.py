import os
import argparse as argp
# Librerias para extraer metadata
from docx import Document
import openpyxl as px
import PyPDF2 as pdf

def wordMetadata(ruta):
    doc = Document(ruta)
    metadata = {}
    prop = doc.core_properties
    metadata["author"] = prop.author
    metadata["category"] = prop.category
    metadata["comments"] = prop.comments
    metadata["content_status"] = prop.content_status
    metadata["created"] = prop.created
    metadata["identifier"] = prop.identifier
    metadata["keywords"] = prop.keywords
    metadata["last_modified_by"] = prop.last_modified_by
    metadata["language"] = prop.language
    metadata["modified"] = prop.modified
    metadata["subject"] = prop.subject
    metadata["title"] = prop.title
    metadata["version"] = prop.version
    return metadata

def pdfMetadata(ruta):
    pdf_reader = pdf.PdfReader(ruta)
    metadata = pdf_reader.metadata
    return metadata

def excelMetadata(ruta):
    wb = px.load_workbook(ruta)
    metadata = {}
    metadata = dict(wb.properties.__dict__)
    
    return metadata
    
def main():

    # Variables de conteo
    wordC = 0
    pdfC = 0
    excelC = 0
    typeF = ''

    # Argumentos por consola
    parser = argp.ArgumentParser(description="Recibir carpeta")

    parser.add_argument("-f", '--folder', help="Folder a analizar")
    opts = parser.parse_args()
    print('Folder a analizar: ', opts.folder)
    
    # Exploración de archivos
    for root, dirs, files in os.walk(opts.folder):
        for file in files:
            ruta = os.path.join(root,file)
            # Explora archivos word
            if file.endswith('.docx'):
                metadata = wordMetadata(ruta)
                wordC += 1
                typeF = '.docx'
            # Explora archivos de excel
            elif file.endswith('.xlsx'):
                metadata = excelMetadata(ruta)
                excelC += 1
                typeF = '.xlsx'
            # Explora archivos pdf
            elif file.endswith('.pdf'):
                metadata = pdfMetadata(ruta)
                pdfC += 1
                typeF = '.pdf'
            else:
                continue # Evitar que se rompa el código
            print(f'Metadata del archivo con extensión {typeF} con ruta {ruta} = \n')
            for key, value in metadata.items():
                print(f'{key}:{value}')
            #print(metadata)
            print('---------------------------------')
            #print('Siguiente archivo...')
    
    print(f'La cuenta total de archivos en {opts.folder} es: ')
    print(f'''
          Archivos word = {wordC} en total \n
          Archivos pdf = {pdfC} en total \n
          Archivos excel = {excelC} en total 
          ''')
            
if __name__ == '__main__':
    main()      
        
