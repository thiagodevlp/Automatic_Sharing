
import pandas as pd
import gspread
import datetime 
consultor = input("nome consultor: ")

def passarcli():

    #apresentando credencial e chamando a spreadsheet.
    gc = gspread.service_account('CREDENCIAL/project-f5-411718-e91bf68189c8.json')
    #planilha de origem
    sh = gc.open_by_url('https://docs.google.com/spreadsheets/d/1M8v86-gIrIGnaVrLR9iFAM2hLjZarKmdsOaq5iPX6Hc/edit#gid=1856095954')
    #Obtendo a aba 0 da spreadsheet.
    worksheet = sh.get_worksheet(0)

    planilhas_destino = {
    "VINICIUS": "https://docs.google.com/spreadsheets/d/1AkOq7fZk3Tv463w-zTEZGp-pyo3Irj8Io3Oi1wEDH5M/edit#gid=462160606",
    "MILENA": "https://docs.google.com/spreadsheets/d/1ryS2J5R4Ga0inzlSRqX0Riq0Qz1E7gUAJh0e4hyXhWA/edit#gid=15516538",
    "MATHEUS": "https://docs.google.com/spreadsheets/d/14lruGzkHIYg9D8yG8plmlrbXR3RvyT4Jcl9Ysvo0YZw/edit#gid=1918432480",
    "LARISSA": "https://docs.google.com/spreadsheets/d/1rqEADE09fQoJBOMoN8xO2zGehmOLF3KmQx_1C8BDPXk/edit#gid=74158966",
    }
    nome_consultor = (consultor)
    if nome_consultor not in planilhas_destino:
        print(f"Nome '{nome_consultor}' não encontrado. Verifique a lista de nomes válidos.")
        exit()

    url_planilha_destino = planilhas_destino[nome_consultor]


    #planilha de destino
    sh2 = gc.open_by_url(url_planilha_destino)    #Obtendo a aba 0 da spreadsheet2.
    worksheet2 = sh2.get_worksheet(0)
    
    qtd_clientes = int(input("Quantidade de clientes? "))
    def filtrar_linhas(worksheet, consultor):
        linhas_selecionadas = []  # Garante que a variável esteja definida
        for linha in worksheet.get_all_values():
            if linha[12] == 'DISPONIVEL' and linha[13] == 'BL' or 'HIBRIDO/END DIFERENTE' in linha[13] and linha[11] != consultor:
                linhas_selecionadas.append(linha)
        return linhas_selecionadas

    #Faz uma seleção 
    linhas_selecionadas = filtrar_linhas(worksheet, consultor)

    # Seleciona 70 linhas aleatórias
    df = pd.DataFrame.from_records(linhas_selecionadas).sample(qtd_clientes, random_state=42)

  
    #Seleciona as colunas 1 a 5 da seleção
    df = df[df.columns[0:6]]


    #insere uma coluna na primeira posição do df com o nome DATA e a data atual como seu conteúdo
    df.insert(0, '0', datetime.datetime.today().strftime('%d/%m/%Y'))

    #Selelciona as colunas com a adição da coluna adiciona [0:7]
    df = df[df.columns[0:7]]


    #worksheet2 recebe a as colunas secionads com a quantidade informada de linhas.
    worksheet2.append_rows(df.values.tolist())
    
    #Apaga a coluna 0 adicionada anteriormente.
    df = df.drop(columns="0")
    
    return df

