import pandas as pd
from os import path

def check_generate_main_config_sheet(filepath):
    print(filepath)
    if not path.isfile(filepath):
        #create excel file with periods configuration sheet
        writer = pd.ExcelWriter(filepath, engine='xlsxwriter')
        workbook = writer.book
        worksheet = workbook.add_worksheet('Configuração_turnos')

        fmt = workbook.add_format({'bold': 1, 'border' : 1, 'align': 'center_across'})
        worksheet.write(0, 0, 'Turno', fmt)
        worksheet.write(0, 1, 'Turmas', fmt)
        worksheet.write(0, 7, 'Aulas por dia', fmt)

        writer.save()

        print("Planilha de configuração geral criada. Por favor configure os turnos e turmas")
        exit(0)
    else:
        sheets = pd.ExcelFile(filepath).sheet_names
        turnos = {}
        max_turmas = 0
        if(len(sheets) == 1):
            #create other sheets based on period configuration
            salas_df = pd.read_excel(filepath, index_col=0, sheet_name='Configuração_turnos', skiprows=1, header=None)
            salas_df = clean_df(salas_df)
            salas_df.fillna('', inplace=True)

            salas = []
            for index, turno in enumerate(salas_df.index):
                filtrado = [s for s in salas_df.loc[turno].tolist()[:-1] if s != '']
                salas.append(filtrado)
                turnos[turno] = filtrado
                max_turmas = max(max_turmas, len(filtrado))

            writer = pd.ExcelWriter(filepath, engine='xlsxwriter')
            workbook = writer.book
            fmt = workbook.add_format({'bold': 1, 'border' : 1, 'align': 'center_across'})
            worksheet = workbook.add_worksheet('Configuração_turnos')
            worksheet.write(0, 0, 'Turno', fmt)
            worksheet.write(0, 1, 'Turmas', fmt)
            worksheet.write(0, max_turmas + 1, 'Aulas por dia', fmt)

            for row_index, (index, row) in enumerate(salas_df.iterrows()):
                worksheet.write(row_index + 1, 0, index)
                for col_index, item in enumerate(list(row)):
                    worksheet.write(row_index + 1, col_index + 1, item)

            similar_sheets = ['Aulas_por_prof', 'Materias_por_prof', 'Itinerários_por_prof', 'Limites_diarios']
            border_format = workbook.add_format({'border': 1})

            for sheet_name in similar_sheets:
                worksheet = workbook.add_worksheet(sheet_name)
                
                if sheet_name == 'Limites_diarios':
                    worksheet.write(1, 0, 'PROFESSOR', fmt)
                    for i, dia in enumerate(['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta']):
                        worksheet.write(0, i + 1, dia, fmt)
                        worksheet.write(1, i + 1, 0, border_format)

                else:
                    if sheet_name == 'Aulas_por_prof':
                        worksheet.write(2, 0, 'DIAS_SEM_ULTIMA_AULA', fmt)
                    else:
                        worksheet.write(2, 0, 'PROFESSOR', fmt)

                    next_start_col = 1
                    for nome_turno, salas_turno in turnos.items():
                        worksheet.merge_range(0, next_start_col, 0, next_start_col + len(salas_turno) - 1, nome_turno, fmt)
                        
                        for i, sala in enumerate(salas_turno):
                            worksheet.write(1, next_start_col + i, sala, fmt)
                            worksheet.write(2, next_start_col + i, 0, border_format)
                        
                        next_start_col = next_start_col + len(salas_turno)
            
            relations_worksheet = workbook.add_worksheet('Relações_itinerários')
            relations_worksheet.write(0, 0, 'Turma 1', fmt)
            relations_worksheet.write(0, 1, 'Turma 2', fmt)

            writer.save()

def clean_df(dataframe):
    try:
        dataframe = dataframe.loc[: , ~dataframe.columns.str.contains('Unnamed')]
        dataframe.dropna(inplace=True)
    except:
        pass
    return dataframe

def check_generate_restrictions_and_constants_spreadsheets(config_directory, turnos):
    aulas_por_dia_por_turno = [entry[1] for entry in turnos.values()]
    max_aulas_por_dia = max(aulas_por_dia_por_turno)

    stop_flag = check_generate_restrictions_spreadsheet(path.join(config_directory, 'Restricoes.xlsx'), max_aulas_por_dia, turnos)
    stop_flag = stop_flag | check_generate_constants_spreadsheet(path.join(config_directory, 'Constantes.xlsx'), max_aulas_por_dia, turnos)
    if stop_flag:
        print("Por favor preencha as planihas de configuração criadas e reinicie o programa")
        exit(0)

def check_generate_restrictions_spreadsheet(filepath, aulas_por_dia, turnos):
    if not path.isfile(filepath):
        iterables = [['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta'], list(range(aulas_por_dia))]
        index = pd.MultiIndex.from_product(iterables, names=["Dia", "Aula"])

        writer = pd.ExcelWriter(filepath, engine='xlsxwriter')
        workbook = writer.book
        worksheet = workbook.add_worksheet("Nome do professor")

        start_col = 1
        start_row = 0

        first_row_format = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'bottom': 2})
        for i, name in enumerate(index.names):
            worksheet.write(start_row, i, name, first_row_format)
        
        i = 0
        for salas, aulas_por_dia_turno in turnos.values():
            for col in salas:
                worksheet.write(start_row, start_col + i + 1, col, first_row_format)
                i+=1
        
        dia_format = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'top': 2, 'bottom': 2})
        for i, dia in enumerate(iterables[0]):
            worksheet.merge_range(i * aulas_por_dia + 1, 0, (i + 1) * aulas_por_dia, 0, dia, dia_format)

        for row_index, aula in enumerate(iterables[1] * 5):
            aula_fmt = workbook.add_format({'bold': True, 'border': 1, 'align': 'center_across'})
            border_fmt = workbook.add_format({'border':1})
            no_class_fmt = workbook.add_format({'border':1, 'bg_color': 'black'})

            if(row_index % aulas_por_dia == aulas_por_dia - 1):
                aula_fmt.set_bottom(2)
                border_fmt.set_bottom(2)
                no_class_fmt.set_bottom(2)

            worksheet.write(row_index + start_row + 1, start_col, aula + 1, aula_fmt)
            
            i = 0
            for salas, aulas_por_dia_turno in turnos.values():
                for sala in salas:
                    if(aula + 1 > aulas_por_dia_turno):
                        worksheet.write_blank(row_index + start_row + 1, start_col + i + 1, None, no_class_fmt)
                    else:
                        worksheet.write_blank(row_index + start_row + 1, start_col + i + 1, None, border_fmt)
                    
                    i+=1

        writer.save()
        print("Planilha de configuração de restrições criada")
        return True
    
    return False

def check_generate_constants_spreadsheet(filepath, aulas_por_dia, turnos):
    if not path.isfile(filepath):
        iterables = [['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta'], list(range(aulas_por_dia))]
        index = pd.MultiIndex.from_product(iterables, names=["Dia", "Aula"])

        writer = pd.ExcelWriter(filepath, engine='xlsxwriter')
        workbook = writer.book
        worksheet = workbook.add_worksheet()

        start_col = 1
        start_row = 0

        first_row_format = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'bottom': 2})
        for i, name in enumerate(index.names):
            worksheet.write(start_row, i, name, first_row_format)
        
        i = 0
        for salas, aulas_por_dia_turno in turnos.values():
            for col in salas:
                worksheet.write(start_row, start_col + i + 1, col, first_row_format)
                i+=1
        
        dia_format = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'top': 2, 'bottom': 2})
        for i, dia in enumerate(iterables[0]):
            worksheet.merge_range(i * aulas_por_dia + 1, 0, (i + 1) * aulas_por_dia, 0, dia, dia_format)

        aula_fmt = workbook.add_format({'bold': True, 'border': 1, 'align': 'center_across'})
        border_fmt = workbook.add_format({'border':1})
        no_class_fmt = workbook.add_format({'border':1, 'bg_color': 'black'})

        for row_index, aula in enumerate(iterables[1] * 5):
            aula_fmt = workbook.add_format({'bold': True, 'border': 1, 'align': 'center_across'})
            border_fmt = workbook.add_format({'border':1})
            no_class_fmt = workbook.add_format({'border':1, 'bg_color': 'black'})

            if(row_index % aulas_por_dia == aulas_por_dia - 1):
                aula_fmt.set_bottom(2)
                border_fmt.set_bottom(2)
                no_class_fmt.set_bottom(2)

            worksheet.write(row_index + start_row + 1, start_col, aula + 1, aula_fmt)

            i = 0
            for salas, aulas_por_dia_turno in turnos.values():
                for sala in salas:
                    if(aula + 1 > aulas_por_dia_turno):
                        worksheet.write_blank(row_index + start_row + 1, start_col + i + 1, None, no_class_fmt)
                    else:
                        worksheet.write_blank(row_index + start_row + 1, start_col + i + 1, None, border_fmt)
                    
                    i+=1

        writer.save()
        print("Planilha de configuração de constantes criada")
        return True

    return False