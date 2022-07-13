import pandas as pd
import os
from os import path
import string
import matplotlib.pyplot as plt
from matplotlib.colors import hsv_to_rgb
import random
from sys import argv

from sheets_generator import clean_df

if(len(argv) < 2):
    print("Por favor especifique um diretório de configuração")
    exit(0)

directory = argv[1]
config_path = path.join(directory, 'Aulas por Professor.xlsx')
sheets = pd.ExcelFile(config_path).sheet_names

salas_df = pd.read_excel(config_path, index_col=0, sheet_name='Configuração_turnos', skiprows=1, header=None)
salas_df.fillna('', inplace=True)

salas_por_periodo = {}
aulas_por_dia = {}
todas_salas = []
for turno in salas_df.index:
    filtrado = [s for s in salas_df.loc[turno].tolist()[:-1] if s != '']
    salas_por_periodo[turno] = filtrado
    aulas_por_dia[turno] = salas_df.loc[turno].tolist()[-1]
    todas_salas += filtrado

converter_dict = {}

for s in todas_salas:
    converter_dict[s] = int

iterables = [['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta'], list(range(1, max(aulas_por_dia.values()) + 1, 1))]
index = pd.MultiIndex.from_product(iterables, names=["Dia", "Aula"])

#Leitura das planilhas de configuração
df_aulas = pd.read_excel(config_path, index_col=0, sheet_name='Aulas_por_prof', converters=converter_dict, nrows=50)
df_aulas.name = 'Aulas por professor'
df_aulas = clean_df(df_aulas)
df_aulas = df_aulas[todas_salas]

# df_materias_por_prof = pd.read_excel(config_path, index_col=0, sheet_name='Materias_por_prof', converters=converter_dict, nrows=50)
# df_materias_por_prof.name = 'Matérias por professor'
# df_materias_por_prof = clean_df(df_materias_por_prof)
# df_materias_por_prof = df_materias_por_prof[todas_salas]

if 'Limites_diarios' in sheets:
    df_limites_diarios = pd.read_excel(config_path, index_col=0, sheet_name='Limites_diarios', nrows=50, converters=converter_dict)
    df_limites_diarios = clean_df(df_limites_diarios)
    df_limites_diarios.to_csv(path.join(directory, 'limites_diarios.csv'), index=False, header=False)

# restrictions = pd.read_excel(path.join(directory, 'Restricoes.xlsx'), index_col=[0,1], sheet_name=None, nrows=max_aulas_por_dia*5)

# constants_df = pd.read_excel(path.join(directory, 'Constantes.xlsx'), index_col=[0,1], nrows=max_aulas_por_dia*5)
# constants_df = constants_df.loc[: , ~constants_df.columns.str.contains('Unnamed')]
# constants_df.fillna('', inplace=True)

if 'Itinerários_por_prof' in sheets:
    itinerarios_df = pd.read_excel(config_path, index_col=0, sheet_name='Itinerários_por_prof', nrows=50, converters=converter_dict)
    itinerarios_df = clean_df(itinerarios_df)

    itinerarios_relations_df = pd.read_excel(config_path, index_col=None, header=None, sheet_name='Relações_Itinerários', usecols=list(range(2)), nrows=50, skiprows=1)
    itinerarios_relations_df = clean_df(itinerarios_relations_df)

    sala_correspondente = {}
    for _, row in itinerarios_relations_df.iterrows():
        r = row.tolist()
        sala_correspondente[r[0]] = r[1]
        sala_correspondente[r[1]] = r[0]

nomes = list(df_aulas.index)[1:]
n_professores = len(nomes)
step = 1 / (n_professores)

def rgb_to_hex(rgb):
    return "#{:02X}{:02X}{:02X}".format(*rgb)

formats = {}
for i, nome in enumerate(nomes):
    h = step * i
    s = random.uniform(0.8, 1)
    v = random.uniform(0.8, 1)

    rgb = hsv_to_rgb((h, s, v))
    rgb = [int(c * 255) for c in rgb]
    formats[nome] = rgb_to_hex(rgb)

def is_lonely_class(dia, a):
    if(a % len(dia) != len(dia) - 1):
        if(dia[a] == dia[a + 1]):
            return 0
    if(a % len(dia) != 0):
        if(dia[a] == dia[a - 1]):
            return 0
    return 1

alg_output_dir = os.path.join(directory, 'outputs')
for filename in os.listdir(alg_output_dir):
    if ".txt" not in filename:
        continue

    cost = filename.split('_')[0]

    dfs_por_periodo = {}

    df = pd.read_csv(os.path.join(alg_output_dir, filename), names=todas_salas, index_col=False)
    df.fillna('', inplace=True)
    df.index = index

    horario_filename = os.path.join(alg_output_dir, 'horario_{}.xlsx'.format(cost))
    writer = pd.ExcelWriter(horario_filename, engine='xlsxwriter')
    for periodo, salas in salas_por_periodo.items():
        df_turno = df[salas]
        n_aulas_atual = len(df_turno.index.levels[1])
        for i in range(aulas_por_dia[periodo] + 1, n_aulas_atual + 1, 1):
            df_turno = df_turno.drop(index=i, level=1)
        
        professores_do_periodo = [professor for professor in nomes if df_aulas.loc[professor, salas].sum() > 0]
        duplas_por_professor = pd.DataFrame(index=professores_do_periodo, columns=salas)
        duplas_por_professor.fillna(0, inplace = True)

        df_counts = pd.DataFrame(columns=nomes, index=df_turno.index)
        df_counts.fillna(0, inplace = True)

        dfs_por_periodo[periodo] = df_turno
        #dfs_por_periodo[periodo].to_excel(writer, sheet_name=periodo)

        workbook = writer.book
        worksheet = workbook.add_worksheet(periodo)

        start_col = 1
        start_row = 0

        first_row_format = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'bottom': 2})
        for i, name in enumerate(df_turno.index.names):
            worksheet.write(start_row, i, name, first_row_format)
        
        for i, col in enumerate(df_turno.columns):
            worksheet.write(start_row, start_col + i + 1, col, first_row_format)
        
        dia_format = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'top': 2, 'bottom': 2})
        for i, dia in enumerate(iterables[0]):
            worksheet.merge_range(i * aulas_por_dia[periodo] + 1, 0, (i + 1) * aulas_por_dia[periodo], 0, dia, dia_format)

        col_widths = []
        for column in df_turno:
            col_widths.append(max(df[column].astype(str).map(len).max(), len(column)))
        
        worksheet.set_column(0, 0, max(col_widths) + 1)
        worksheet.set_column(2, len(df_turno.columns) + len(df_turno.index.levels) - 1, max(col_widths) + 1)
            
        
        for row_index, (aula, row) in enumerate(dfs_por_periodo[periodo].iterrows()):
            for col_index, (sala, professor) in enumerate(row.iteritems()):
                fmt = workbook.add_format({'bold': True, 'right': 1, 'left': 1})
                if(professor != ''):
                    fmt.set_color(formats[professor])
                    if 'Itinerários_por_prof' in sheets:
                        if(itinerarios_df.loc[professor, sala] > 0):
                            professor_correspondente = dfs_por_periodo[periodo].loc[aula, sala_correspondente[sala]]
                            if(itinerarios_df.loc[professor_correspondente, sala_correspondente[sala]] > 0):
                                fmt.set_bg_color('yellow')
                
                aula_fmt = workbook.add_format({'bold': True, 'border': 1, 'align': 'center_across'})
                if(row_index % aulas_por_dia[periodo] == aulas_por_dia[periodo] - 1):
                    fmt.set_bottom(2)
                    aula_fmt.set_bottom(2)

                worksheet.write(row_index + start_row + 1, start_col, aula[1], aula_fmt)
                worksheet.write(row_index + start_row + 1, col_index + start_col + 1, professor, fmt)
                
                if(professor != ''):
                    df_counts.loc[aula, professor] += 1

        for col_index, sala in enumerate(df_turno.columns):
            for dia in iterables[0]:
                aulas = df_turno.loc[dia, sala]
                counts = aulas.value_counts()
                professores = list(counts.index)
                aulas_list = list(aulas)
                for professor in professores:
                    if professor == '':
                        continue
                    lonely_classes = 0
                    for i in range(aulas_por_dia[periodo]):
                        if(aulas_list[i] == professor):
                            lonely_classes += is_lonely_class(aulas_list, i)
                    if((counts.loc[professor] == 2 and lonely_classes == 0) or (counts.loc[professor] == 3 and lonely_classes != 0)):
                        duplas_por_professor.loc[professor, sala] += 1
        duplas_por_professor.sort_index(inplace=True)

        duplas_por_professor_filtered = pd.DataFrame(columns=duplas_por_professor.columns)
        for professor, row in duplas_por_professor.iterrows():
            for sala, n_duplas in row.iteritems():
                if n_duplas == 0 and df_aulas.loc[professor, sala] > 1:
                    duplas_por_professor_filtered.loc[professor] = duplas_por_professor.loc[professor]

        df_janelas = pd.DataFrame(columns = iterables[0], index=nomes)
        df_janelas.fillna(0, inplace=True)
        
        for dia in iterables[0]:
            for professor in nomes:
                counts = df_counts.loc[dia, professor].value_counts()
                n_janelas = 0
                if(0 in counts.index and 1 in counts.index):
                    if(counts[1] > 1):
                        aulas = df_counts.loc[dia, professor].tolist()
                        aulas_no_dia = len(aulas)
                        vistos = []
                        for i in range(aulas_no_dia):
                            if aulas[i] == 0 and i not in vistos:
                                found_top = False
                                found_bottom = False
                                for j in range(i - 1, -1, -1):                                        
                                    if aulas[j] == 1:
                                        found_top = True
                                        break
                                    else:
                                        vistos.append(j)
                                for j in range(i + 1, aulas_no_dia, 1):
                                    if aulas[j] == 1:
                                        found_bottom = True
                                        break
                                    else:
                                        vistos.append(j)
                                if found_top and found_bottom:
                                    df_janelas.loc[professor, dia] += 1
        
        df_janelas = df_janelas[df_janelas.sum(axis=1).gt(0)]
        df_janelas.sort_index(inplace=True)

        #Inserção de estatisticas de janelas na planilha de horários
        statistics_row = aulas_por_dia[periodo] * 5 + 2
        statistics_col = 1

        title_fmt = workbook.add_format({'bold': True, 'align': 'center_across', 'border': 1})
        border_fmt = workbook.add_format({'align': 'center_across', 'border': 1})
        problem_fmt = workbook.add_format({'align': 'center_across', 'border': 1, 'bg_color': 'red'})

        if(len(df_janelas) > 0):
            n_cols = len(df_janelas.columns)
            worksheet.merge_range(statistics_row, statistics_col, statistics_row, statistics_col + n_cols, 'Janelas', title_fmt)
            statistics_row += 1

            worksheet.write(statistics_row, statistics_col, None, border_fmt)
            for i, dia in enumerate(df_janelas.columns):
                worksheet.write(statistics_row, statistics_col + 1 + i, dia, title_fmt)
            statistics_row += 1

            for professor, row in df_janelas.iterrows():
                worksheet.write(statistics_row, statistics_col, professor, title_fmt)
                for i, n_janelas in enumerate(list(row)):
                    worksheet.write(statistics_row, statistics_col + 1 + i, n_janelas, border_fmt)
                statistics_row += 1
            statistics_row += 1

        #Inserção de estatisticas de aulas duplas
        n_cols = len(df_turno.columns)
        if(len(duplas_por_professor_filtered > 0)):
            worksheet.merge_range(statistics_row, statistics_col, statistics_row, statistics_col + n_cols, 'Aulas duplas - Problemas', title_fmt)
            statistics_row += 1

            worksheet.write(statistics_row, statistics_col, None, border_fmt)
            for i, sala in enumerate(duplas_por_professor_filtered.columns):
                worksheet.write(statistics_row, statistics_col + 1 + i, sala, title_fmt)
            statistics_row += 1

            for professor, row in duplas_por_professor_filtered.iterrows():
                worksheet.write(statistics_row, statistics_col, professor, title_fmt)
                for i, (sala, n_duplas) in enumerate(row.iteritems()):
                    if n_duplas == 0 and df_aulas.loc[professor, sala] > 1:
                        worksheet.write(statistics_row, statistics_col + 1 + i, n_duplas, problem_fmt)
                    else:
                        worksheet.write(statistics_row, statistics_col + 1 + i, n_duplas, border_fmt)
                statistics_row += 1
            statistics_row += 1
        
        worksheet.merge_range(statistics_row, statistics_col, statistics_row, statistics_col + n_cols, 'Aulas duplas', title_fmt)
        statistics_row += 1

        worksheet.write(statistics_row, statistics_col, None, border_fmt)
        for i, sala in enumerate(duplas_por_professor.columns):
            worksheet.write(statistics_row, statistics_col + 1 + i, sala, title_fmt)
        statistics_row += 1

        for professor, row in duplas_por_professor.iterrows():
            worksheet.write(statistics_row, statistics_col, professor, title_fmt)
            for i, n_duplas in enumerate(list(row)):
                worksheet.write(statistics_row, statistics_col + 1 + i, n_duplas, border_fmt)
            statistics_row += 1
        statistics_row += 1


        df_alternativo = pd.DataFrame(columns=index, index=professores_do_periodo)
        df_alternativo.fillna(' ', inplace=True)

        for aula, row in dfs_por_periodo[periodo].iterrows():
            for sala, professor in row.iteritems():
                if(professor != ''):
                    df_alternativo.loc[professor, aula] = sala
        
        df_alternativo.sort_index(inplace=True)
        
        worksheet = workbook.add_worksheet(periodo + '-Alternativo')

        statistics_row = 0
        statistics_col = 0

        worksheet.write(statistics_row, statistics_col, 'Dia', title_fmt)
        for i, dia in enumerate(iterables[0]):
            start_col = statistics_col + 1 + i * aulas_por_dia[periodo]
            end_col = statistics_col + (i + 1) * aulas_por_dia[periodo]
            worksheet.merge_range(statistics_row, start_col, statistics_row, end_col, dia, title_fmt)
        statistics_row += 1

        worksheet.write(statistics_row, statistics_col, 'Aula', title_fmt)
        for col_index, aula in enumerate(list(range(1, aulas_por_dia[periodo] + 1, 1)) * 5):
            worksheet.write(statistics_row, statistics_col + 1 + col_index, aula, title_fmt)
        statistics_row += 1

        for professor, row in df_alternativo.iterrows():
            worksheet.write(statistics_row, statistics_col, professor, title_fmt)
            i = 0
            for aula, sala in row.iteritems():
                if aula[1] > aulas_por_dia[periodo]:
                    continue

                border_fmt = workbook.add_format({'align': 'center_across', 'border': 1})
                if 'Itinerários_por_prof' in sheets and sala != ' ':
                    if(itinerarios_df.loc[professor, sala] > 0):
                        professor_correspondente = df_turno.loc[aula, sala_correspondente[sala]]
                        if(itinerarios_df.loc[professor_correspondente, sala_correspondente[sala]] > 0):
                            border_fmt.set_bg_color('yellow')

                if aula[1] == 1:
                    border_fmt.set_left(2)
                worksheet.write(statistics_row, statistics_col + 1 + i, sala, border_fmt)

                i+=1
            statistics_row += 1
        statistics_row += 1
        

        # for i, aula in enumerate(df_alternativo.columns):
        #     worksheet.write(statistics_row, statistics_col + 1 + i, dia, title_fmt)
        # statistics_row += 1

        # for professor, row in df_janelas.iterrows():
        #     worksheet.write(statistics_row, statistics_col, professor, title_fmt)
        #     for i, n_janelas in enumerate(list(row)):
        #         worksheet.write(statistics_row, statistics_col + 1 + i, n_janelas, border_fmt)
        #     statistics_row += 1
        # statistics_row += 1
        

    writer.save()


#  for val, color in formats.items():
#             fmt = workbook.add_format({'color': color, 'bold': True})
#             worksheet.conditional_format('A1:Z100', {'type': 'cell',
#                                                 'criteria': '=',
#                                                 'value': val,
#                                                 'format': fmt})