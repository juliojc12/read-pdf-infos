from PyPDF2 import PdfReader
import openpyxl
import re
import os
import tkinter as tk
from tkinter import filedialog
from colorama import Fore, Back, Style

root = tk.Tk()
root.withdraw()

dir = filedialog.askdirectory()
#print(dir)

diretorio = dir

workbook = openpyxl.Workbook()
sheet = workbook.active

sheet['A1'] = 'NOME'
sheet['B1'] = 'CPF'
sheet['C1'] = 'CÓDIGO BARRAS'
sheet['D1'] = 'VALOR'
sheet['E1'] = 'DATA VENCIMENTO'
sheet['F1'] = 'ANUIDADE'
sheet['G1'] = 'NÚMERO BOLETO'

linha = 2

for nome_arquivo in os.listdir(diretorio):
    extension = os.path.splitext(nome_arquivo)[1]
    if extension == ".pdf":
        try:
            caminho_arquivo = os.path.join(diretorio, nome_arquivo)
            if os.path.isfile(caminho_arquivo):
                reader = PdfReader(caminho_arquivo)
                pages = reader.pages[0]
                text = pages.extract_text((0, 90))

                cpf_regex = re.compile(r'\d{3}\.\d{3}\.\d{3}\-\d{2}')
                boleto = re.compile(r'\d{5}\.\d{5} \d{5}\.\d{6} \d{5}\.\d{6} \d \d{14}')
                anuidade_regex = re.compile(r'\d\ \/\ \d{2}')
                numero_boleto_regex = re.compile(r'\d{12} \d')


                numero_boleto = numero_boleto_regex.search(text)
                cpf_search = cpf_regex.search(text)
                bar_code = boleto.search(text)
                cod_barra = bar_code.group()
                cpf = cpf_search.group()


                linhas = text.splitlines();
                central = linhas[29]

                if central == "SAC CAIXA : 0800 726 0101 (informações, reclamações, sugestões e elogios)":
                    nome = linhas[36].split('-')[0]
                    anuidade_linha = linhas[5]
                    data = linhas[34].split(" ")[0]
                    valor = linhas[35]
                else:
                    nome = linhas[30].split('-')[0]
                    anuidade_linha = linhas[5]
                    data = linhas[28].split(" ")[0]
                    valor = linhas[29]

                anuidade = anuidade_regex.search(anuidade_linha).group()
                boleto = numero_boleto.group().split(" ")[0]
                print(f"{Fore.GREEN}{Back.BLACK}Salvando: {nome}{Fore.RESET}{Back.RESET}")
                sheet[f'A{linha}'] = nome
                sheet[f'B{linha}'] = cpf
                sheet[f'C{linha}'] = cod_barra
                sheet[f'D{linha}'] = valor
                sheet[f'E{linha}'] = data
                sheet[f'F{linha}'] = anuidade
                sheet[f'G{linha}'] = boleto

                linha += 1
        except:
            pass #not a good pratice, but... sorry ;(
    else:
        continue

desktop_dir = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
workbook.save(f'{diretorio}/PagamentosOAB.xlsx')
print(f"{Back.BLACK}{Fore.YELLOW}\n\nPlanilha salva em: {diretorio}/PagamentosOAB.xlsx {Back.RESET}{Fore.RESET}")