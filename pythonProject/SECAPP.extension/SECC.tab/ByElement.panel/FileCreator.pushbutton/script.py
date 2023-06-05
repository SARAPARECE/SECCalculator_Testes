#! python3
# -*- coding: utf-8 -*-
__title__   = "File Creator"
__author__  = "Sara Parece"
__doc__     = """This button creates the necessary documents for the calculator"""

import clr
import Autodesk.Revit.DB as DB
import sys
import os
diretorio_base = os.path.dirname(os.path.abspath(__file__))
diretorio_base1 = os.path.dirname(os.path.abspath(diretorio_base))
diretorio_base2 = os.path.dirname(os.path.abspath(diretorio_base1))
diretorio_base3 = os.path.dirname(os.path.abspath(diretorio_base2))
diretorio_base4 = os.path.dirname(os.path.abspath(diretorio_base3))
diretorio_base5 = os.path.dirname(os.path.abspath(diretorio_base4))
diretorio_base6 = os.path.dirname(os.path.abspath(diretorio_base5))
diretorio_pacote = os.path.join(diretorio_base6, 'anaconda3', 'envs', 'PyRevit', 'Lib', 'site-packages')
sys.path.append(diretorio_pacote)

#sys.path.append(r'C:\Users\Asus\anaconda3\envs\PyRevit\Lib\site-packages')
import pandas as pd
import json
import csv





clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

doc = __revit__.ActiveUIDocument.Document

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
        #type_name = element_type.ToDSType(False).Name
        type_name = element.Name
        #DB.ElementType.Name.GetValue(element)
        volume = 0

        try:
            volume= element.get_Parameter(DB.BuiltInParameter.HOST_VOLUME_COMPUTED).AsDouble() * 0.0283168 #HOST_VOLUME_COMPUTED convert cubic feet to cubic meters
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
            "ElementID": element.Id,
            "SECClasS_Code": shared_parameter_value,
            "SECClasS_Title": shared_parameter_value2,
            "Family and Type": "{} - {}".format(family_name, type_name),
            "Volume": round(volume, 2),
            "Area": round(area, 2),
            "Length": round(length, 2),
            "Phase_Created": phase_parameter_value2
        }
        data.append(Export_BIM_data)




####### SEC-WBS #######
script_dir = os.path.dirname(__file__) # path do script a correr
parentDirectory = os.path.dirname(script_dir) # path pai do script
SEC_WBS_file_path = os.path.join(parentDirectory, 'FileCreator.pushbutton', 'SEC_WBS.csv') # path da pasta e nome do cdv a abrir
with open(SEC_WBS_file_path, mode='r', encoding='utf-8-sig') as SEC_WBS:
    SEC_WBS_table = csv.DictReader(SEC_WBS, delimiter=";")
    SEC_WBS_data = list(SEC_WBS_table)
#print(type(SEC_WBS_data))
#print(SEC_WBS_data)
#SEC_WBS_json = json.dumps(SEC_WBS_data, ensure_ascii=False).encode('utf8')


###### Tabela excel ######
Co2Value_file_path = os.path.join(parentDirectory, 'FileCreator.pushbutton', 'Co2Value.csv')
with open(Co2Value_file_path, mode='r', encoding='utf-8-sig') as SEC_WBS:
    Co2Value_table = csv.DictReader(SEC_WBS, delimiter=";")
    Co2Value_data = list(Co2Value_table)
#Co2Value_json = Co2Value_table.to_json(orient='records', force_ascii=False).encode('utf8')
#Co2Value_data = json.loads(Co2Value_json)  # Extrat data from JSON

## Create List CSV - Building Information ##
csv_file_path1 = os.path.join(os.path.dirname(__file__), "Building_Information.csv")
add_line =["*Mandatory", "", "", "", ""]
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

#2 MAPEAMENTO SECCLASS WBS E MEDIDAS
for i in range(len(data)):
  for row in SEC_WBS_data:
    if data[i]['SECClasS_Code'] is None:
        pass
    else:
        if data[i]['SECClasS_Code'] == row['SECClasS_Code']:
            data[i]['Quantity of elements'] = 0
            data[i]['WBS_L1'] = row['WBS_L1']
            data[i]['WBS_Title_L1'] = row['WBS_Title_L1']
            data[i]['WBS_L2'] = row['WBS_L2']
            data[i]['WBS_Title_L2'] = row['WBS_Title_L2']
            data[i]['WBS_L3'] = row['WBS_L3']
            data[i]['WBS_Title_L3'] = row['WBS_Title_L3']
            data[i]['WBS_Code'] = row['WBS_Code']
            data[i]['Expected_lifespan'] = row['Expected_lifespan']
            data[i]['Mass'] = None
            data[i]['Co2_Total'] = None
            data[i]['BLC_Mass_Total'] = None
            data[i]['BLC_Co2_Total'] = None
            data[i]['Normalised requirement factor over building lifetime'] = None
            data[i]['Measure'] = row['Measure']
            data[i]['Conversion_Factor'] = row['Conversion Factor (Kg/m3, Kg/m2;kg/m;k/U)']
            data[i]['Unit_Cost'] = row['Unit_Cost']
            data[i]['Cost'] = None
            data[i]['GWP_A1-A3'] = row['GWP A1-A3 (Kg/m3, Kg/m2;kg/m;k/U)']
            data[i]['Concrete (%)'] = row['Concrete (%)']
            data[i]['Bricks (%)'] = row['Bricks (%)']
            data[i]['Tiles (%)'] = row['Tiles (%)']
            data[i]['Ceramics (%)'] = row['Ceramics (%)']
            data[i]['Wood (%)'] = row['Wood (%)']
            data[i]['Glass (%)'] = row['Glass (%)']
            data[i]['Plastic (%)'] = row['Plastic (%)']
            data[i]['Bituminous mixtures (%)'] = row['Bituminous mixtures (%)']
            data[i]['Copper/bronze/brass (%)'] = row['Copper/bronze/brass (%)']
            data[i]['Aluminium (%)'] = row['Aluminium (%)']
            data[i]['Iron/steel (%)'] = row['Iron/steel (%)']
            data[i]['Other metal (%)'] = row['Other metal (%)']
            data[i]['Soil and stones (%)'] = row['Soil and stones (%)']
            data[i]['Dredging spoil (%)'] = row['Dredging spoil (%)']
            data[i]['Track ballast (%)'] = row['Track ballast (%)']
            data[i]['Insulation materials (%)'] = row['Insulation materials (%)']
            data[i]['Asbestos containing materials (%)'] = row['Asbestos containing materials (%)']
            data[i]['Gypsum-based materials (%)'] = row['Gypsum-based materials (%)']
            data[i]['Electrical and Electronic Equipment (%)'] = row['Electrical and Electronic Equipment (%)']
            data[i]['Cables (%)'] = row['Cables (%)']


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
    if 'Measure' in row:
        # create a temporary dictionary
        temp_dict = {
            'SECClasS_Code': row['SECClasS_Code'],
            'SECClasS_Title': row['SECClasS_Title'],
            'Conversion_Factor': row['Conversion_Factor'],
            'Quantity of elements': row['Quantity of elements'],
            'Unit_Cost': row['Unit_Cost'],
            'GWP_A1-A3': row['GWP_A1-A3'],
            'Concrete (%)': row['Concrete (%)'],
            'Bricks (%)': row['Bricks (%)'],
            'Tiles (%)': row['Tiles (%)'],
            'Ceramics (%)': row['Ceramics (%)'],
            'Wood (%)': row['Wood (%)'],
            'Glass (%)': row['Glass (%)'],
            'Plastic (%)': row['Plastic (%)'],
            'Bituminous mixtures (%)': row['Bituminous mixtures (%)'],
            'Copper/bronze/brass (%)': row['Copper/bronze/brass (%)'],
            'Aluminium (%)': row['Aluminium (%)'],
            'Iron/steel (%)': row['Iron/steel (%)'],
            'Other metal (%)': row['Other metal (%)'],
            'Soil and stones (%)': row['Soil and stones (%)'],
            'Dredging spoil (%)': row['Dredging spoil (%)'],
            'Track ballast (%)': row['Track ballast (%)'],
            'Insulation materials (%)': row['Insulation materials (%)'],
            'Asbestos containing materials (%)': row['Asbestos containing materials (%)'],
            'Gypsum-based materials (%)': row['Gypsum-based materials (%)'],
            'Electrical and Electronic Equipment (%)': row['Electrical and Electronic Equipment (%)'],
            'Cables (%)': row['Cables (%)'],
        }

        # add to the appropriate dictionary and update seen_codes
        if row['Measure'] == 'V':
            if temp_dict['SECClasS_Code'] not in seen_codes:
                list_V[len(list_V)] = temp_dict
                seen_codes.add(temp_dict['SECClasS_Code'])
        elif row['Measure'] == 'A':
            if temp_dict['SECClasS_Code'] not in seen_codes:
                list_A[len(list_A)] = temp_dict
                seen_codes.add(temp_dict['SECClasS_Code'])
        elif row['Measure'] == 'L':
            if temp_dict['SECClasS_Code'] not in seen_codes:
                list_L[len(list_L)] = temp_dict
                seen_codes.add(temp_dict['SECClasS_Code'])
        elif row['Measure'] == 'U':
            if temp_dict['SECClasS_Code'] not in seen_codes:
                list_U[len(list_U)] = temp_dict
                seen_codes.add(temp_dict['SECClasS_Code'])
        else:
            print("Error: Tou have codes with out a unit of measure, please confirm in SEC-WBS.xlxs", len(list_V) + len(list_A) + len(list_L) + len(list_U))
    else:
        list_empty[len(list_empty)] = row

# count the SECClasS_Code values
num_values_V = {}
num_values_A = {}
num_values_L = {}
num_values_U = {}
for lst, num_values in [(list_V, num_values_V), (list_A, num_values_A), (list_L, num_values_L), (list_U, num_values_U)]:
    counter = [item['SECClasS_Code'] for item in lst.values()]
    num_values.update({code: count for code, count in zip(set(counter), map(counter.count, set(counter)))})



# TESTE
#print(len(data))
#print(num_values_A)
#print(len(num_values_V))
#print(len(num_values_L))
#print(len(num_values_U))
#print(len(list_empty))
#print(len(list_V_temp) + len(list_A_temp) + len(list_L_temp) + len(list_U_temp) + len(list_empty))


## Create CSV - Criar ficheiro UNICO utilizadores ##
add_line = ["Warning", "These density values are reference values that can be edited. All values must be filled.", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]
csv_file_path = os.path.join(os.path.dirname(__file__), "Building_Elements_Information.csv")
with open(csv_file_path, mode='w', newline='', encoding='utf-8') as csv_file:
    CSV = csv.writer(csv_file)
    CSV.writerow(["ref",
          "SECClasS_Code",
          "SECClasS_Title",
          "Quantity of elements",
          "Measure",
          "Conversion Factor (kg/m3, kg/m2, kg/m, kg/u)",
          "GWP A1-A3 (kgCo2e/m3, kgCo2e/m2, kgCo2e/m, kgCo2e/u)",
          "Unit Cost (€/unit of measure)",
          "Concrete (%)",
          "Bricks (%)",
          "Tiles (%)",
          "Ceramics (%)",
          "Wood (%)",
          "Glass (%)",
          "Plastic (%)",
          "Bituminous mixtures (%)",
          "Copper/bronze/brass (%)",
          "Aluminium (%)",
          "Iron/steel (%)",
          "Other metal (%)",
          "Soil and stones (%)",
          "Dredging spoil (%)",
          "Track ballast (%)",
          "Insulation materials (%)",
          "Asbestos containing materials (%)",
          "Gypsum-based materials (%)",
          "Electrical and Electronic Equipment (%)",
          "Cables (%)",
          "Total (must be 100%)"])
    for new_k, new_v in list_V.items():
      for keys, value in num_values_V.items():
        if new_v['SECClasS_Code'] is keys:
          CSV.writerow([new_k, new_v['SECClasS_Code'], new_v['SECClasS_Title'], new_v['Quantity of elements'], "V", new_v['Conversion_Factor'], new_v['GWP_A1-A3'], new_v['Unit_Cost'], new_v['Concrete (%)'], new_v['Bricks (%)'], new_v['Ceramics (%)'], new_v['Wood (%)'], new_v['Glass (%)'], new_v['Plastic (%)'], new_v['Bituminous mixtures (%)'], new_v['Copper/bronze/brass (%)'], new_v['Aluminium (%)'], new_v['Iron/steel (%)'], new_v['Other metal (%)'], new_v['Soil and stones (%)'], new_v['Dredging spoil (%)'], new_v['Track ballast (%)'], new_v['Track ballast (%)'], new_v['Insulation materials (%)'], new_v['Asbestos containing materials (%)'], new_v['Gypsum-based materials (%)'], new_v['Electrical and Electronic Equipment (%)'], new_v['Cables (%)'], "0"])
    for new_k, new_v in list_A.items():
      for keys, value in num_values_A.items():
        if new_v['SECClasS_Code'] is keys:
          CSV.writerow([new_k, new_v['SECClasS_Code'], new_v['SECClasS_Title'], new_v['Quantity of elements'], "A", new_v['Conversion_Factor'], new_v['GWP_A1-A3'], new_v['Unit_Cost'], new_v['Concrete (%)'], new_v['Bricks (%)'], new_v['Ceramics (%)'], new_v['Wood (%)'], new_v['Glass (%)'], new_v['Plastic (%)'], new_v['Bituminous mixtures (%)'], new_v['Copper/bronze/brass (%)'], new_v['Aluminium (%)'], new_v['Iron/steel (%)'], new_v['Other metal (%)'], new_v['Soil and stones (%)'], new_v['Dredging spoil (%)'], new_v['Track ballast (%)'], new_v['Track ballast (%)'], new_v['Insulation materials (%)'], new_v['Asbestos containing materials (%)'], new_v['Gypsum-based materials (%)'], new_v['Electrical and Electronic Equipment (%)'], new_v['Cables (%)'], "0"])
    for new_k, new_v in list_L.items():
      for keys, value in num_values_L.items():
        if new_v['SECClasS_Code'] is keys:
          CSV.writerow([new_k, new_v['SECClasS_Code'], new_v['SECClasS_Title'], new_v['Quantity of elements'], "L", new_v['Conversion_Factor'], new_v['GWP_A1-A3'], new_v['Unit_Cost'], new_v['Concrete (%)'], new_v['Bricks (%)'], new_v['Ceramics (%)'], new_v['Wood (%)'], new_v['Glass (%)'], new_v['Plastic (%)'], new_v['Bituminous mixtures (%)'], new_v['Copper/bronze/brass (%)'], new_v['Aluminium (%)'], new_v['Iron/steel (%)'], new_v['Other metal (%)'], new_v['Soil and stones (%)'], new_v['Dredging spoil (%)'], new_v['Track ballast (%)'], new_v['Track ballast (%)'], new_v['Insulation materials (%)'], new_v['Asbestos containing materials (%)'], new_v['Gypsum-based materials (%)'], new_v['Electrical and Electronic Equipment (%)'], new_v['Cables (%)'], "0"])
    for new_k, new_v in list_U.items():
      for keys, value in num_values_U.items():
        if new_v['SECClasS_Code'] is keys:
          CSV.writerow([new_k, new_v['SECClasS_Code'], new_v['SECClasS_Title'], new_v['Quantity of elements'], "U", new_v['Conversion_Factor'], new_v['GWP_A1-A3'], new_v['Unit_Cost'], new_v['Concrete (%)'], new_v['Bricks (%)'], new_v['Ceramics (%)'], new_v['Wood (%)'], new_v['Glass (%)'], new_v['Plastic (%)'], new_v['Bituminous mixtures (%)'], new_v['Copper/bronze/brass (%)'], new_v['Aluminium (%)'], new_v['Iron/steel (%)'], new_v['Other metal (%)'], new_v['Soil and stones (%)'], new_v['Dredging spoil (%)'], new_v['Track ballast (%)'], new_v['Track ballast (%)'], new_v['Insulation materials (%)'], new_v['Asbestos containing materials (%)'], new_v['Gypsum-based materials (%)'], new_v['Electrical and Electronic Equipment (%)'], new_v['Cables (%)'], "0"])
    CSV.writerow(add_line)
    csv_file.close()


# Define a separator line
separator = "_______________________________________________________________________________________________________________________________"

# Define a header for the program
header = "........................................................SECCALCULATOR.........................................................."
header_centered = header.center(len(separator))



# Print the formatted text
print(separator)
print(header_centered)
print(separator)
print("The CSV files were created: Building_Information.csv and Building_Elements_Information.csv.")
print("The files will be stored in the FileCreator.pushbutton folder.")
print("Both documents must be filled in with the correct information.")
print(separator)



