import os
import pandas as pd
import glob
import openpyxl
from openpyxl.styles import Font, NamedStyle, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

def get_excel_files(user_name: str) -> list:
    all_files = os.listdir(os.path.join("C:\\Users","Marcos","Desktop", "Meses"))
    excel_files = []
    for file in all_files:
        if file.endswith(".xlsx"):
            excel_files.append("C:/Users/Marcos/Desktop/Meses/" + file)
    return excel_files

def nota_de_credito(data_nueva: pd.DataFrame) -> pd.DataFrame:
    for i in range(len(data_nueva)):
        if "Nota de Crédito" in data_nueva.iloc[i,1]:
            data_nueva.iloc[i,4] = data_nueva.iloc[i,4]*-1
    return data_nueva

def data_cronologica(data: pd.DataFrame) -> pd.DataFrame:
    data['Fecha'] = pd.to_datetime(data['Fecha'], format='%d/%m/%Y')
    data = data.sort_values(by='Fecha', ascending=True).reset_index(drop=True)
    data['Fecha'] = data['Fecha'].dt.strftime('%d/%m/%Y') 
    return data

def main_menu():
    while True:
        print("Menu Principal:\n1) Actualizar Excel de Cliente\n2) Crear Excel Para Nuevo Cliente")
        choice = input("Ingrese un Número: ")
        if choice == "1":
            os.system('cls')
            print("Has Seleccionado la Opcion 1. Usted Va a Actualizar los Excels de los Siguientes Clientes:")
            cuits = pd.read_excel('C:/Users/Marcos/Desktop/CUITS.xlsx')
            file_paths = get_excel_files(1)
            clientes_cuits = []
            for c in range(len(cuits)):
                for file_path in file_paths:
                    if str(cuits.iloc[c,1]) in file_path:
                        clientes_cuits.append(cuits.iloc[c,0])
            clientes_cuits = list(set(clientes_cuits))
            for cliente in clientes_cuits:
                print(cliente)
            print("\nDesea continuar?\n1)Si\n2)No, volver al menú")
            eleccion=input("Ingrese un Número: ")
            if eleccion == "1":
                actualizacion_excel()
                os.system('cls')
                print("Los Excels de los siguientes clientes han sido actualizados:")
                for cliente in clientes_cuits:
                    print(cliente)
                asd = input("\nVolver al menú?\n1)Si\n2)No, salir del programa\nIngrese un Número: ")
                if asd =="1":
                    main_menu()
                else:
                    break
            else:
                main_menu()
        elif choice == "2":
            main_menu()
        else:
            print("Opción invalida")
            main_menu()

def actualizacion_excel():
    cuits = pd.read_excel('C:/Users/Marcos/Desktop/CUITS.xlsx')
    file_paths = get_excel_files(1)
    i=0
    for c in range(len(cuits)):
        a=0
        if i != len(file_paths):
            alldata_emitidos = pd.DataFrame()
            alldata_recibidos = pd.DataFrame()
            for file_path in file_paths:
                if str(cuits.iloc[c,1]) in file_path:
                    if 'Emitidos' in file_path:
                        alldata_emitidos = pd.concat([alldata_emitidos, pd.read_excel(file_path,skiprows=1)], axis=0, ignore_index=True).reset_index(drop=True)
                    else:
                        alldata_recibidos = pd.concat([alldata_recibidos, pd.read_excel(file_path,skiprows=1)], axis=0, ignore_index=True).reset_index(drop=True)  
                    i+=1
            if len(alldata_emitidos) > 0:
                data_nueva = data_cronologica(alldata_emitidos[['Fecha', 'Tipo', 'Número Desde', 'Denominación Receptor', 'Imp. Total']])
                data_nueva = nota_de_credito(data_nueva)
                excel_filepath = f'C:/Users/Marcos/Desktop/Oficina/Monotributo/{cuits.iloc[c, 0]}.xlsx'
                data_excel = pd.read_excel(excel_filepath, sheet_name="VENTAS NUEVO", usecols="A:E", skiprows=5)
                size = len(data_excel)
                wb = openpyxl.load_workbook(excel_filepath)
                sheet = wb['VENTAS NUEVO']
                for row_idx, row in enumerate(openpyxl.utils.dataframe.dataframe_to_rows(data_nueva, index=False, header=False), (size+7)):
                    for col_idx, value in enumerate(row, 1):
                        sheet.cell(row=row_idx, column=col_idx, value=value)
                for row in wb["VENTAS NUEVO"].iter_rows(min_row=(size+7), max_row=(size+len(data_nueva)+7), min_col=1, max_col=5):
                    for cell in row:
                        cell.font = Font(name="Calibri", size=11) 
                        if cell.column == 1:
                            cell.alignment = Alignment(horizontal="right")  
                        elif cell.column == 3:
                            cell.alignment = Alignment(horizontal="center")  
                wb.save(excel_filepath)
            if len(alldata_recibidos) > 0:
                data_nueva = data_cronologica(alldata_recibidos[['Fecha', 'Tipo', 'Número Desde', 'Denominación Emisor', 'Imp. Total']])
                data_nueva = nota_de_credito(data_nueva)
                excel_filepath = f'C:/Users/Marcos/Desktop/Oficina/Monotributo/{cuits.iloc[c, 0]}.xlsx'
                data_excel = pd.read_excel(excel_filepath, sheet_name="COMPRAS NUEVO", usecols="A:E", skiprows=5)
                size = len(data_excel)
                wb = openpyxl.load_workbook(excel_filepath)
                sheet = wb['COMPRAS NUEVO']
                for row_idx, row in enumerate(openpyxl.utils.dataframe.dataframe_to_rows(data_nueva, index=False, header=False), (size+7)):
                    for col_idx, value in enumerate(row, 1):
                        sheet.cell(row=row_idx, column=col_idx, value=value)
                for row in wb["COMPRAS NUEVO"].iter_rows(min_row=(size+7), max_row=(size+len(data_nueva)+7), min_col=1, max_col=5):
                    for cell in row:
                        cell.font = Font(name="Calibri", size=11) 
                        if cell.column == 1:
                            cell.alignment = Alignment(horizontal="right")  
                wb.save(excel_filepath)
        else:
            break

# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -



if __name__ == "__main__":
    main_menu()