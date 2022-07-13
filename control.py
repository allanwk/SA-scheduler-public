import pandas as pd
import csv
import numpy as np
from sys import argv
from os import path

from sheets_generator import clean_df, \
                            check_generate_main_config_sheet, \
                            check_generate_restrictions_and_constants_spreadsheets

if(len(argv) < 2):
    print("Por favor especifique um diretório de configuração")
    exit(0)
    
directory = argv[1]
config_path = path.join(directory, 'Aulas por Professor.xlsx')

check_generate_main_config_sheet(config_path)

sheets = pd.ExcelFile(config_path).sheet_names

salas_df = pd.read_excel(config_path, index_col=0, sheet_name='Configuração_turnos', skiprows=1, header=None)
salas_df.fillna('', inplace=True)

salas = []
period_starts = [0]
aulas_por_dia_por_turno = []
salas_por_turno = []
turnos = {}

for index, turno in enumerate(salas_df.index):
    filtrado = [s for s in salas_df.loc[turno].tolist()[:-1] if s != '']
    salas.append(filtrado)
    aulas_por_dia_por_turno.append(salas_df.loc[turno].tolist()[-1])
    salas_por_turno.append(len(filtrado))
    turnos[turno] = (filtrado, salas_df.loc[turno].tolist()[-1])
    if(index != len(salas_df.index) - 1):
        period_starts.append(period_starts[-1] + len(filtrado))

with open(path.join(directory, 'periods.csv'), 'w', encoding='UTF8', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(period_starts)
    writer.writerow(aulas_por_dia_por_turno)
    writer.writerow(salas_por_turno)

todas_salas = []
for arr in salas:
    todas_salas += arr

def get_turno(sala):
    for index, turno in enumerate(salas):
        if(sala in turno):
            return index

max_aulas_por_dia = max(aulas_por_dia_por_turno)
check_generate_restrictions_and_constants_spreadsheets(directory, turnos)

converter_dict = {}

for turno in salas:
    for s in turno:
        converter_dict[s] = int

#Leitura das planilhas de configuração
df_aulas = pd.read_excel(config_path, index_col=0, sheet_name='Aulas_por_prof', converters=converter_dict, nrows=50)
df_aulas.name = 'Aulas por professor'
df_aulas = clean_df(df_aulas)
df_aulas = df_aulas[todas_salas]

df_materias_por_prof = pd.read_excel(config_path, index_col=0, sheet_name='Materias_por_prof', converters=converter_dict, nrows=50)
df_materias_por_prof.name = 'Matérias por professor'
df_materias_por_prof = clean_df(df_materias_por_prof)
df_materias_por_prof = df_materias_por_prof[todas_salas]

if 'Limites_diarios' in sheets:
    df_limites_diarios = pd.read_excel(config_path, index_col=0, sheet_name='Limites_diarios', nrows=50, converters=converter_dict)
    df_limites_diarios = clean_df(df_limites_diarios)
    df_limites_diarios.to_csv(path.join(directory, 'limites_diarios.csv'), index=False, header=False)

restrictions = pd.read_excel(path.join(directory, 'Restricoes.xlsx'), index_col=[0,1], sheet_name=None, nrows=max_aulas_por_dia*5)

constants_df = pd.read_excel(path.join(directory, 'Constantes.xlsx'), index_col=[0,1], nrows=max_aulas_por_dia*5)
constants_df = constants_df.loc[: , ~constants_df.columns.str.contains('Unnamed')]
constants_df.fillna('', inplace=True)

if 'Itinerários_por_prof' in sheets:
    itinerarios_df = pd.read_excel(config_path, index_col=0, sheet_name='Itinerários_por_prof', nrows=50, converters=converter_dict)
    itinerarios_df = clean_df(itinerarios_df)
    itinerarios_df.to_csv(path.join(directory, 'itinerarios.csv'), index=False, header=False)

    itinerarios_relations_df = pd.read_excel(config_path, index_col=None, header=None, sheet_name='Relações_Itinerários', usecols=list(range(2)), nrows=50, skiprows=1)
    itinerarios_relations_df = clean_df(itinerarios_relations_df)

    for index, sala in enumerate(todas_salas):
        itinerarios_relations_df = itinerarios_relations_df.replace(sala, index)

    itinerarios_relations_df.to_csv(path.join(directory, 'itinerarios_relations.csv'), index=False, header=False)

nomes = list(df_aulas.index)[1:]
dias = ['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta']
restriction_csv_rows = []
constants_csv_rows = []

#Tratamento de restrições e constantes
for professor, restrictions_df in restrictions.items():
    if(professor not in nomes):
        if(professor == 'Nome do professor'):
            continue
        raise Exception("Nome de professor {} passado nas restrições não encontrado.".format(professor))
    restrictions_df = restrictions_df.loc[: , ~restrictions_df.columns.str.contains('Unnamed')]
    restrictions_df.fillna('', inplace=True)

    for index, row in restrictions_df.iterrows():
        for sala, restricao in row.iteritems():
            dia, aula = index
            dia = dias.index(dia)
            aula -= 1
            if 'X' in restricao:
                peso = restricao.split('X')[-1]
                sala_index = list(restrictions_df.columns).index(sala)
                restriction_csv_rows.append([sala_index, dia, aula, nomes.index(professor), peso, 0])
            if 'U' in restricao:
                peso = restricao.split('U')[-1]
                sala_index = list(restrictions_df.columns).index(sala)
                restriction_csv_rows.append([sala_index, dia, aula, nomes.index(professor), peso, 1])

for index, row in constants_df.iterrows():
    for sala, professor in row.iteritems():
        if professor == '':
            continue

        if professor == 'X':
            dia, aula = index
            dia = dias.index(dia)
            aula -= 1
            sala_index = list(constants_df.columns).index(sala)
            constants_csv_rows.append([sala_index, dia, aula, -1])

        elif professor in nomes:
            dia, aula = index
            dia = dias.index(dia)
            aula -= 1
            sala_index = list(constants_df.columns).index(sala)
            constants_csv_rows.append([sala_index, dia, aula, nomes.index(professor)])
        else:
            raise Exception('Professor {} na planilha de constantes não encontrado.'.format(professor))

#Geração dos arquivos csv
df_aulas.to_csv(path.join(directory, 'aulas_por_sala.csv'), index=True, header=False)
df_materias_por_prof.to_csv(path.join(directory, 'materias_por_professor.csv'), index=False, header=False)

with open(path.join(directory, 'restrictions.csv'), 'w', encoding='UTF8', newline='') as file:
    writer = csv.writer(file)
    writer.writerows(restriction_csv_rows)

with open(path.join(directory, 'constants.csv'), 'w', encoding='UTF8', newline='') as file:
    writer = csv.writer(file)
    writer.writerows(constants_csv_rows)
