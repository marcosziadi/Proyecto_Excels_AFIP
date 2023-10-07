import os
import pandas as pd
import glob
import openpyxl
from openpyxl.styles import Font, NamedStyle, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
import shutil
import warnings
warnings.filterwarnings("ignore", category=pd.errors.SettingWithCopyWarning)
warnings.filterwarnings("ignore", category=FutureWarning)

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
        os.system('cls')
        print("Menu Principal:\n1) Actualizar Excel de Cliente\n2) Crear Excel Para Nuevo Cliente\n3) Salir del programa")
        choice = input("Ingrese un Número: ")
        if choice == "1":
            os.system('cls')
            print("Has Seleccionado la Opcion 'Actualizar Excel de Cliente'. Usted Va a Actualizar los Excels de los Siguientes Clientes:")
            cuits = pd.read_excel('C:/Users/Marcos/Desktop/PRUEBA CUITS.xlsx')
            file_paths = get_excel_files(1)
            clientes_cuits = []
            for c in range(len(cuits)):
                for file_path in file_paths:
                    if str(cuits.iloc[c,1]) in file_path:
                        clientes_cuits.append(cuits.iloc[c,0])
            clientes_cuits = list(set(clientes_cuits))
            for cliente in clientes_cuits:
                print(cliente)
            print("\nDesea continuar?\n1)Si, actualizarlos\n2)No, volver al menú")
            eleccion=input("Ingrese un Número: ")
            if eleccion == "1":
                actualizacion_excel()
                os.system('cls')
                print("Los Excels de los siguientes clientes han sido actualizados:")
                for cliente in clientes_cuits:
                    print(cliente)
                asd = input("\nVolver al menú?\n1)Si\n2)No, salir del programa\nIngrese un Número: ")
                if asd !="1":
                    break
            else:
                os.system('cls')
        elif choice == "2":
            os.system('cls')
            print("Has Seleccionado la Opcion 'Crear Excel Para Nuevo Cliente'.")
            nuevo_excel()
            os.system('cls')
            abc = input("El excel del cliente ha sido creado. Volver al menú?\n1)Si\n2)No, salir del programa\nIngrese un Número: ")
            if abc != "1":
                break
        else:
            break

def actualizacion_excel():
    cuits = pd.read_excel('C:/Users/Marcos/Desktop/PRUEBA CUITS.xlsx')
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

def nuevo_excel():
    nene=0
    while True:
        cuits = pd.read_excel('C:/Users/Marcos/Desktop/PRUEBA CUITS.xlsx')
        cuits2 = cuits
        nuevo_cliente = pd.Series([input("Ingrese el Nombre del Nuevo Cliente: ").upper(),
                                input("Ingrese el CUIT del Nuevo Cliente: ")],
                                index=cuits.columns)
        cuits2 = cuits2.append(nuevo_cliente, ignore_index=True)
        eleccion = input(f"\nUsted está por crear un Excel para el cliente {cuits2.iloc[len(cuits2)-1,0]} y su CUIT es {cuits2.iloc[len(cuits2)-1,1]}\nQue desea hacer?\n1)Crear Excel\n2)Cambiar el Nombre o el CUIT\n3)Volver al menú\nIngrese un Número:")
        if eleccion == "1":
            cuits = cuits2
            cuits.to_excel('C:/Users/Marcos/Desktop/PRUEBA CUITS.xlsx', index=False)
            break
        elif eleccion == "3":
            nene=1
            break
        else:
            os.system('cls')
    if nene == 0:
        file_paths = get_excel_files(1)
        alldata_emitidos = pd.DataFrame()
        alldata_recibidos = pd.DataFrame()

        for file_path in file_paths:
            if str(cuits.iloc[(len(cuits)-1),1]) in file_path:
                if 'Emitidos' in file_path:
                    alldata_emitidos = pd.concat([alldata_emitidos, pd.read_excel(file_path,skiprows=1)], axis=0, ignore_index=True).reset_index(drop=True)
                else:
                    alldata_recibidos = pd.concat([alldata_recibidos, pd.read_excel(file_path,skiprows=1)], axis=0, ignore_index=True).reset_index(drop=True)  
        source_file = "C:/Users/Marcos/Desktop/Oficina/Monotributo/modelo.xlsx"
        excel_file_path = f"C:/Users/Marcos/Desktop/Oficina/Monotributo/{(cuits.iloc[(len(cuits)-1),0]).upper()}.xlsx"
        shutil.copyfile(source_file, excel_file_path)


        if len(alldata_emitidos) > 0:
            data_ventas = data_cronologica(alldata_emitidos[['Fecha', 'Tipo', 'Número Desde', 'Denominación Receptor', 'Imp. Total']])
            data_ventas = nota_de_credito(data_ventas)
            data_excel = pd.read_excel(excel_file_path, sheet_name="VENTAS NUEVO", usecols="A:E", skiprows=5)
            size = len(data_excel)
            wb = openpyxl.load_workbook(excel_file_path)
            sheet = wb['VENTAS NUEVO']
            for row_idx, row in enumerate(openpyxl.utils.dataframe.dataframe_to_rows(data_ventas, index=False, header=False), (size+7)):
                for col_idx, value in enumerate(row, 1):
                    sheet.cell(row=row_idx, column=col_idx, value=value)
            for row in wb["VENTAS NUEVO"].iter_rows(min_row=(size+7), max_row=(size+len(data_ventas)+7), min_col=1, max_col=5):
                for cell in row:
                    cell.font = Font(name="Calibri", size=11) 
                    if cell.column == 1:
                        cell.alignment = Alignment(horizontal="right")  
                    elif cell.column == 3:
                        cell.alignment = Alignment(horizontal="center")  
            wb.save(excel_file_path)
        if len(alldata_recibidos) > 0:
            data_compras = data_cronologica(alldata_recibidos[['Fecha', 'Tipo', 'Número Desde', 'Denominación Emisor', 'Imp. Total']])
            data_compras = nota_de_credito(data_compras)
            data_excel = pd.read_excel(excel_file_path, sheet_name="COMPRAS NUEVO", usecols="A:E", skiprows=5)
            size = len(data_excel)
            wb = openpyxl.load_workbook(excel_file_path)
            sheet = wb['COMPRAS NUEVO']
            for row_idx, row in enumerate(openpyxl.utils.dataframe.dataframe_to_rows(data_compras, index=False, header=False), (size+7)):
                for col_idx, value in enumerate(row, 1):
                    sheet.cell(row=row_idx, column=col_idx, value=value)
            for row in wb["COMPRAS NUEVO"].iter_rows(min_row=(size+7), max_row=(size+len(data_compras)+7), min_col=1, max_col=5):
                for cell in row:
                    cell.font = Font(name="Calibri", size=11) 
                    if cell.column == 1:
                        cell.alignment = Alignment(horizontal="right")  
            wb.save(excel_file_path)
# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

if __name__ == "__main__":
    main_menu()