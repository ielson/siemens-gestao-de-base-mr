import pandas as pd
import datetime
import argparse
from openpyxl import load_workbook

parser = argparse.ArgumentParser(description='Gera xls de peças')
parser.add_argument('data', type=str, help='a data da planilha a ser usada')
args = parser.parse_args()



print("Iniciando o programa")
data_atual = datetime.date.today()
if data_atual.month == 2 and data_atual.day == 29:
    data_atual = data_atual.replace(day=28)
data_limite = data_atual.replace(year=2018)
print("Data de vencimento das Tales: {}".format(data_limite))

data_planilha = args.data
print("Data da planilha: {}".format(data_planilha))
base_instalada_path = "{}/Base Instalada.xlsx".format(data_planilha)
consumiveis_path = "{}/Consumíveis MR.xlsx".format(data_planilha)

base_instalada = pd.read_excel(base_instalada_path, sheet_name='Planilha3')
print("Tamanho da base instalada: {}".format(base_instalada.shape[0]))
consumiveis = pd.read_excel(consumiveis_path, sheet_name='Planilha1')

base_instalada_tales = base_instalada[base_instalada['TALES'] == 'Y']
print("Dessas {} tem tales".format(base_instalada_tales.shape[0]))
consumiveis_tales = consumiveis[(consumiveis['Material'] == 10018247) | (consumiveis['Material'] == 7391886)]

consumiveis_tales_novas = consumiveis_tales[consumiveis_tales['Dt.Consumo'].dt.date > data_limite]
print("Das Tales substituidas {} estão no prazo".format(consumiveis_tales_novas.shape[0]))

consumiveis_tales_novas_destination = consumiveis_tales_novas[(consumiveis_tales_novas['Destinação Mat.'] == 'CB - CONSUMED BILLABLE') |
                                                  (consumiveis_tales_novas['Destinação Mat.'] == 'C - CONSUMED')]
print("Das Tales novas {} foram consumidas".format(consumiveis_tales_novas_destination.shape[0]))

#teste = base_instalada_tales[consumiveis_tales_novas_destination['Nº série'] in base_instalada_tales['Nº de série']]
base_tales_vencida = base_instalada_tales[~base_instalada_tales['Nº de série'].isin(consumiveis_tales_novas_destination['Nº série'])]
print("Número de Tales Vencidas: {}".format(base_tales_vencida.shape[0]))

numero_por_estado = base_tales_vencida.groupby(['Rg'])['TALES'].count()
numero_por_estado['Total'] = base_tales_vencida.shape[0]
numero_por_estado_data = numero_por_estado.rename(data_planilha)
#teste.to_excel('{}/resultados.xlsx'.format(data_planilha))

planilha = load_workbook('{}/resultados.xlsx'.format(data_planilha))
writer = pd.ExcelWriter('{}/resultados.xlsx'.format(data_planilha), engine='openpyxl')
writer.book = planilha
writer.sheets = {ws.title: ws for ws in planilha.worksheets}

for sheetname in writer.sheets:
    numero_por_estado_data.to_excel(writer, sheet_name=sheetname, startcol=writer.sheets[sheetname].max_column, index=False)

writer.save()

