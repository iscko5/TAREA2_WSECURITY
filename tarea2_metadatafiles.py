import os
import argparse as argp
import configparser
# Librerias para extraer metadata
import docx
import openpyxl
import PyPDF2

def main():

    parser = argp.ArgumentParser(description="Recibir carpeta")

    parser.add_argument("-f", '--folder', help="Folder a analizar")
    opts = parser.parse_args()
    print('Folder a analizar: ', opts.folder)
