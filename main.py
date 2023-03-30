from PyPDF2 import PdfReader
import openpyxl
import re
import os
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

dir = filedialog.askdirectory()
print(dir)

diretorio = dir

workbook = openpyxl.Workbook()
sheet = workbook.active

sheet['A1'] = 'NOME'
sheet['B1'] = 'CPF'
sheet['C1'] = 'CÓDIGO BARRAS'
sheet['D1'] = 'VALOR'
sheet['E1'] = 'DATA VENCIMENTO'
sheet['F1'] = 'ANUIDADE'

linha = 2


for nome_arquivo in os.listdir(diretorio):
    caminho_arquivo = os.path.join(diretorio, nome_arquivo)
    if os.path.isfile(caminho_arquivo):
        reader = PdfReader(caminho_arquivo)
        pages = reader.pages[0]
        text = pages.extract_text((0, 90))

        cpf_regex = re.compile(r'\d{3}\.\d{3}\.\d{3}\-\d{2}')
        boleto = re.compile(r'\d{5}\.\d{5} \d{5}\.\d{6} \d{5}\.\d{6} \d \d{14}')
        anuidade_regex = re.compile(r'\d\ \/\ \d{2}')

        cpf_search = cpf_regex.search(text)
        bar_code = boleto.search(text)
        cod_barra = bar_code.group()
        cpf = cpf_search.group()


        linhas = text.splitlines();
        nome = linhas[30].split('-')[0]
        anuidade_linha = linhas[5]
        data = linhas[28].split(" ")[0]
        #anuidade = anuidade_linha[20:30]
        valor = linhas[29]
        anuidade = anuidade_regex.search(anuidade_linha).group()

        sheet[f'A{linha}'] = nome
        sheet[f'B{linha}'] = cpf
        sheet[f'C{linha}'] = cod_barra
        sheet[f'D{linha}'] = valor
        sheet[f'E{linha}'] = data
        sheet[f'F{linha}'] = anuidade

        linha += 1


#print(text)

desktop_dir = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
workbook.save(f'{diretorio}/PagamentosOAB.xlsx')