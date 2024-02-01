import pandas as pd
import gspread
import selecionar_colar
import openpyxl
from selecionar_colar import consultor

#apresentando credencial e chamando a spreadsheet.
gc = gspread.service_account('CREDENCIAL/project-f5-411718-e91bf68189c8.json')
#planilha de origem
sh = gc.open_by_url('https://docs.google.com/spreadsheets/d/1M8v86-gIrIGnaVrLR9iFAM2hLjZarKmdsOaq5iPX6Hc/edit#gid=1856095954')
#Obtendo a aba 0 da spreadsheet.
worksheet = sh.get_worksheet(0) 

"""
#Abrindo planilha excel
wb = openpyxl.load_workbook("PLANILHAS/retorno.xlsx")

# Obtenha a planilha desejada
sheet = wb["planilha01"]
# Obtenha os valores da planilha
values = sheet.values

# Converta os valores em um df
df = pd.DataFrame(values)

"""

# Obtendo todos os dados da planilha, incluindo cabeçalhos
all_data = worksheet.get_all_values()

# Criando o DataFrame
df2 = pd.DataFrame(all_data)

#Chamar função
df = selecionar_colar.passarcli()

# Atualizar o valor da coluna 12 para as linhas que possuem CNPJs presentes no df
df2.loc[df2.iloc[:, 3].isin(df.iloc[:, 3]), 12] = "EM_AÇÃO"
df2.loc[df2.iloc[:, 3].isin(df.iloc[:, 3]), 11] = consultor 

df2.to_excel("PLANILHAS/origem.xlsx")

# Atualizar a planilha Google Sheets
worksheet.update(range_name='', values=df2.values.tolist())