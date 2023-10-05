#! python3
# -*- coding: utf-8 -*-
# Títulos de botões - Só funciona no IronPython
__title__   = "File Creator"
__author__  = "Sara Parece"
__doc__     = """This button creates the necessary documents for the calculator"""

# Importações
import clr
import Autodesk.Revit.DB as DB
from Autodesk.Revit.UI import UIDocument, UIApplication
import sys
import os
import csv
import json
from Autodesk.Revit.UI import TaskDialog
from System.Windows.Forms import FolderBrowserDialog, DialogResult
clr.AddReference("System.Windows.Forms")
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

# Variáveis
doc = __revit__.ActiveUIDocument.Document

# Caixa de dialogo para selecionar caminho dos ficheiros
def select_folder():
    dialog = FolderBrowserDialog()
    dialog.Description = "Select a folder for your files."

    result = dialog.ShowDialog()
    if result == DialogResult.OK:
        selected_folder1 = dialog.SelectedPath
        return selected_folder1
    else:
        return None


def save_selected_folder_to_json(selected_folder):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    parentDirectory = os.path.dirname(script_dir)  # path pai do script
    parentDirectory1 = os.path.dirname(parentDirectory)
    parentDirectory2 = os.path.dirname(parentDirectory1)
    json_file_path = os.path.join(parentDirectory2,'Data','selected_folder.json')

    path = {"selected_folder": selected_folder}

    with open(json_file_path, "w") as json_file:
        json.dump(path, json_file)


if __name__ == "__main__":
    selected_folder = select_folder()
    if selected_folder is not None:
        TaskDialog.Show("Selected Folder", f"Selected folder: {selected_folder}")
        folder = True
        save_selected_folder_to_json(selected_folder)
    else:
        folder = False
        pass

# Extrair info do modelo

if folder == True:
    categories = [
        DB.BuiltInCategory.OST_Walls,
        DB.BuiltInCategory.OST_Floors,
        DB.BuiltInCategory.OST_Roofs,
        DB.BuiltInCategory.OST_StructuralColumns,
        DB.BuiltInCategory.OST_StructuralFraming,
        DB.BuiltInCategory.OST_Columns,
        DB.BuiltInCategory.OST_Doors,
        DB.BuiltInCategory.OST_Windows,
        DB.BuiltInCategory.OST_CurtainWallPanels,
        DB.BuiltInCategory.OST_Stairs,
        DB.BuiltInCategory.OST_StructuralFoundation,
        ]
    data = []

    for category in categories:
        elements = DB.FilteredElementCollector(doc).OfCategory(category).WhereElementIsNotElementType().ToElements()
        for element in elements:
            element_type = doc.GetElement(element.GetTypeId()) # para propriedades
            family_name = element_type.FamilyName
            type_name = element.Name
            volume = 0

            try:
                volume = element.get_Parameter(DB.BuiltInParameter.HOST_VOLUME_COMPUTED).AsDouble() * 0.0283168 #HOST_VOLUME_COMPUTED convert cubic feet to cubic meters
            except:
                pass

            try:
                area = element.get_Parameter(
                    DB.BuiltInParameter.HOST_AREA_COMPUTED).AsDouble() * 0.092903  # convert square feet to square meters
                length = element.get_Parameter(
                    DB.BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 0.3048  # convert feet to meters
            except:
                pass
           #SECCLASS CODES
            shared_parameter = element_type.LookupParameter("ClassificacaoSecclassSsNumero")
            if shared_parameter:
                shared_parameter_value = shared_parameter.AsString()
            else:
                shared_parameter_value = None

            shared_parameter2 = element_type.LookupParameter("ClassificacaoSecclassSsDescricao")
            if shared_parameter2:
                shared_parameter_value2 = shared_parameter2.AsString()
            else:
                shared_parameter_value2 = None

            phase_parameter = element.LookupParameter("Phase Created")
            if phase_parameter:
                phase_parameter_value2 = None
            else:
                phase_parameter_value2 = None
            #classificacao_parameter = element.get_Parameter("04466190-35a1-4404-82d0-898f32b92ec4")
            #classificacao_value = classificacao_parameter.AsString() if classificacao_parameter else None

            Export_BIM_data = {
                "ElementID": str(element.Id),
                "SECClasS_Code": shared_parameter_value,
                "SECClasS_Title": shared_parameter_value2,
                "Family and Type": "{} - {}".format(family_name, type_name),
                "Volume": round(volume, 2),
                "Area": round(area, 2),
                "Length": round(length, 2),
                "Phase_Created": phase_parameter_value2
            }
            data.append(Export_BIM_data)




    # Abrir SEC-WBS #
    script_dir = os.path.dirname(os.path.abspath(__file__))
    parentDirectory = os.path.dirname(script_dir)  # path pai do script
    parentDirectory1 = os.path.dirname(parentDirectory)
    parentDirectory2 = os.path.dirname(parentDirectory1)

    SEC_WBS_file_path = os.path.join(parentDirectory2, 'Data', 'SEC_WBS.csv') # path da pasta e nome do cdv a abrir
    with open(SEC_WBS_file_path, mode='r', encoding='utf-8-sig') as SEC_WBS:
        SEC_WBS_table = csv.DictReader(SEC_WBS, delimiter=";")
        SEC_WBS_data = list(SEC_WBS_table)
        SEC_WBS_dict = dict(SEC_WBS_data[0])
        list_column_names = list(SEC_WBS_dict.keys())

    #2 MAPEAMENTO SECCLASS WBS E MEDIDAS
    for i in range(len(data)):
      for row in SEC_WBS_data:
          for name in list_column_names[2:]:
            if data[i]['SECClasS_Code'] is None:
                pass
            else:
                if data[i]['SECClasS_Code'] == row['SECClasS_Code']:
                    data[i][name] = row[name]
                    data[i]['Quantity of elements'] = 0


    #3 Filtro Phase Created
    list_new = []
    list_existing = []
    for row in data:
      if row['Phase_Created'] == "New Construction":
        list_new.append(row)
      elif row['Phase_Created'] == "Existing":
        list_existing.append(row)
      else:
        list_new.append(row) # ver isto mais tarde
    data = list_new
    #3 Filtro só classificados
    list_classified = []
    list_notclassified = []
    for row in data:
      if row['SECClasS_Code'] is not None:
        list_classified.append(row)
      else:
        list_notclassified.append(row)

    data = list_classified #passa a ser data

    ## add ao Export BIM coluna do id
    for id in range(len(data)):
        data[id]["id"] = id

    #### Contagem de elementos ####
    count_dict = {}

    for row in data:
        if 'SECClasS_Code' in row:
            code = row['SECClasS_Code']
            if code in count_dict:
                count_dict[code] += 1
            else:
                count_dict[code] = 1
    for row in data:
        if 'SECClasS_Code' in row:
            code = row['SECClasS_Code']
            if code in count_dict:
                row['Quantity of elements'] = count_dict[code]

    #############
    # create empty dictionaries and set
    list_V, list_A, list_L, list_U, list_empty = ({}, {}, {}, {}, {})
    seen_codes = set()

    # iterate over data
    for row in data:
        if 'Unit of Measure' in row:

            # add to the appropriate dictionary and update seen_codes
            if row['Unit of Measure'] == 'V':
                if row['SECClasS_Code'] not in seen_codes:
                    list_V[len(list_V)] = row
                    seen_codes.add(row['SECClasS_Code'])
            elif row['Unit of Measure'] == 'A':
                if row['SECClasS_Code'] not in seen_codes:
                    list_A[len(list_A)] = row
                    seen_codes.add(row['SECClasS_Code'])
            elif row['Unit of Measure'] == 'L':
                if row['SECClasS_Code'] not in seen_codes:
                    list_L[len(list_L)] = row
                    seen_codes.add(row['SECClasS_Code'])
            elif row['Unit of Measure'] == 'U':
                if row['SECClasS_Code'] not in seen_codes:
                    list_U[len(list_U)] = row
                    seen_codes.add(row['SECClasS_Code'])
            else:
                message = (
                    "There is SECClass Codes with out a unit of measure, please confirm in SEC-WBS.csv"
                )
                dialog = TaskDialog("SECCALCULATOR Message")
                dialog.MainInstruction = "Error"
                dialog.MainContent = message
                dialog.Show()
        else:
            list_empty[len(list_empty)] = row

    # TESTE
    #print(temp_dict)
    #print(seen_codes)
    #print(list_A)
    #print(len(list_empty))
    #print(len(list_V_temp) + len(list_A_temp) + len(list_L_temp) + len(list_U_temp) + len(list_empty))

    ## Create CSV - Criar ficheiro UNICO utilizadores ##
    add_line = ["Warning", "These values are from SEC-WBS and can be edited. All values must be filled.", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]
    # Função para escrever os dados em um arquivo CSV
    def write_data_to_csv(csv_file, data_list, category):
        for new_k, new_v in data_list.items():
            row = [new_k,new_v['SECClasS_Code'], new_v['SECClasS_Title'], new_v['Quantity of elements'], category,
                     new_v['Conversion Factor (Kg/m3, Kg/m2, kg/m, kg/u)'],
                     new_v['GWP A1-A3 (kgCo2e/m3, kgCo2e/m2, kgCo2e/m, kgCo2e/u)'],
                     new_v['Unit_Cost']]
            for name in list_column_names[15:]:
                row2 = new_v.get(name).replace(';', ',').replace('[', '').replace(']', '')
                #row2 = float(new_v[name]) não dá porque nunca converti , para . por isso é sempre uma string e dá um erro de formato sai assim ["2,5"] para o csv
                row.append(row2)
            row.append("0")
            CSV.writerow(row)

    # Caminho do arquivo CSV
    csv_file_path = os.path.join(selected_folder, "Building_Elements_Information.csv")

    # Dados das listas
    data_lists = {'V': list_V, 'A': list_A, 'L': list_L, 'U': list_U}
    try:
        # Abrir o arquivo CSV
        with open(csv_file_path, mode='w', newline='', encoding='utf-8') as csv_file:
            CSV = csv.writer(csv_file)
            # Escrever cabeçalho
            row_headers = [
                "ref", "SECClasS_Code", "SECClasS_Title", "Quantity of elements", "Unit of Measure",
                "Conversion Factor (kg/m3, kg/m2, kg/m, kg/u)", "GWP A1-A3 (kgCo2e/m3, kgCo2e/m2, kgCo2e/m, kgCo2e/u)",
                "Unit Cost (€/unit of measure)"
            ]
            row_headers.extend(list_column_names[15:])  # Adicione os nomes das colunas restantes
            row_headers.append("Total (must be 100%)")
            CSV.writerow(row_headers)

            # Escrever dados de cada categoria
            for category, data_list in data_lists.items():
                write_data_to_csv(CSV, data_list, category)

            # Escrever a linha de aviso
            CSV.writerow(add_line)

            # Create List CSV - Building Information #
            # csv_file_path1 = os.path.join(os.path.dirname(__file__), "Building_Information.csv")
            csv_file_path1 = os.path.join(selected_folder, "Building_Information.csv")
            add_line = ["*Mandatory", "", "", "", ""]
            with open(csv_file_path1, mode='w', newline='', encoding='utf-8') as list_file_BI:
                CSV = csv.writer(list_file_BI)
                CSV.writerow([
                    "Project Number",
                    "Project Name",
                    "Building Name",
                    "Building GFA (m2)*",
                    "Building lifespan (years)*",
                ])
                CSV.writerow(["", "", "", "0", "0"])
                CSV.writerow(add_line)

        if __name__ == "__main__":
            message = (
                ""
                f"The CSV files were created: Building_Information.csv and Building_Elements_Information.csv.\n"
                ""
                f"The files will be stored in the {selected_folder} folder.\n"
                ""
                f"Both documents must be filled in with the correct information."
            )
            dialog = TaskDialog("SECCALCULATOR Message")
            dialog.MainInstruction = "Operação Concluída"
            dialog.MainContent = message
            dialog.Show()

    except Exception as e:  # Especifica o tipo de exceção
        if __name__ == "__main__":
            message = (
                ""
                f" {e}.\n"
                ""
                f"Try to save and close the files above."
            )
            dialog = TaskDialog("SECCALCULATOR Message")
            dialog.MainInstruction = "SECCalculator - Error"
            dialog.MainContent = message
            dialog.Show()



else:
    if __name__ == "__main__":
        message = (
            "Operation canceled, no folder was selected."
        )
        dialog = TaskDialog("SECCALCULATOR Message")
        dialog.MainInstruction = "Canceled"
        dialog.MainContent = message
        dialog.Show()











