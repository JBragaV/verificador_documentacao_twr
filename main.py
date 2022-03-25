import requests
from bs4 import BeautifulSoup
import os
import openpyxl
from datetime import datetime


def abrir_excel():
    endereco_arquivo = r"D:\projetos\pessoal\sistema_paim\publicacoes_torre_sdco.xlsx"
    return openpyxl.open(endereco_arquivo)


def entrar_pasta(coluna_desejada):
    arquivo = abrir_excel()
    for i in range(1, 5):
        coordenada_coluna = 0
        valor = arquivo.worksheets[1].cell(row=i, column=1).value
        valor = valor.replace("coluna ", "").lower()
        print(valor)
        if valor == coluna_desejada:
            coordenada_coluna_selecionada = arquivo.worksheets[1].cell(row=i, column=2).value
            coordenada_coluna = coordenada_coluna_selecionada
            print(f"Coluna Selecionada {coordenada_coluna}")
            return coordenada_coluna
    path = "S:\Documentação TWR"
    arquivos = os.listdir(path)
    # for arquivo in arquivos:
    #     print(f"{arquivo}")


def converte_data(data):
    # não
    # esquece
    # de
    # implementar
    # isso
    # aqui
    pass

def funcao_teste(regulamento):
    nome_formatado = regulamento.split(" ")
    regulamento = f"{nome_formatado[0]}-{nome_formatado[1]}"
    try:
        r = requests.get(f"https://publicacoes.decea.mil.br/publicacao/{regulamento}")
        codigo_http = r.status_code
        print(codigo_http)
        pagina = BeautifulSoup(r.text, 'html.parser')
        informarcao = pagina.find_all(attrs={"class": "d-flex justify-content-between align-items-center"})
        informarcao = informarcao[0].find_next("strong").text
        print(f"A data de expedição é {informarcao}")
    except:
        nome = regulamento.split("-")
        nome = f"{nome[0]} {nome[1]}-{nome[2]}"
        print(f"A publicação {nome} não foi encotrado")


if __name__ == '__main__':
    numero_coluna = entrar_pasta("data em vigor")
    arquivo = abrir_excel()
    planilha_ativa = arquivo.worksheets[0]
    for i in range(6, 10000):
        valor_celula = planilha_ativa.cell(row=i, column=numero_coluna).value
        if valor_celula is None:
            print("não tem mais valores")
            break
        dia = datetime.date(valor_celula).day
        mes = datetime.date(valor_celula).month
        ano = datetime.date(valor_celula).year
        print(valor_celula)
        print(f"{dia}/{mes}/{ano}")

    # funcao_teste("ica 100-37")
    print('Jocimar Braga')

# Leia o notion para lembrar do fluxo que o programa deve ter
# https://www.notion.so/Verificador-Documenta-es-7c14c607927140ffb5e3792939916a74
