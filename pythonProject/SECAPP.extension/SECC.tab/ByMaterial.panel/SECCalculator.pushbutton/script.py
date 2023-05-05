#! python3
# -*- coding: utf-8 -*-


import clr
import Autodesk.Revit.DB as DB
from Autodesk.Revit.DB import Transaction
import sys
#sys.path.append(r'C:\Users\Asus\anaconda3\envs\PyRevit\Lib\site-packages')
import pandas as pd
import json
import csv
import os

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

def get_material_quantities(element_id):
    element = doc.GetElement(element_id)
    material_quantities = {}
    if element:
        mat_quantities = DB.MaterialQuantities.GetMaterialQuantities(element, DB.UnitTypeId.CubicMeters)
        for mat in mat_quantities:
            material_name = mat.Name
            material_volume = mat.Volume
            if material_name in material_quantities:
                material_quantities[material_name] += material_volume
            else:
                material_quantities[material_name] = material_volume
    return material_quantities

data = []
for category in categories:
    elements = DB.FilteredElementCollector(doc).OfCategory(category).WhereElementIsNotElementType().ToElements()
    for element in elements:
        element_type = doc.GetElement(element.GetTypeId())
        family_name = element_type.FamilyName
        element_Id = element.Id.ToString()
        type_name = element.Name
        category = element.get_Parameter(DB.BuiltInParameter.ELEM_CATEGORY_PARAM_MT).AsValueString()
        volume = 0

        try:
            volume = element.get_Parameter(DB.BuiltInParameter.HOST_VOLUME_COMPUTED).AsDouble() * 0.0283168
        except:
            pass

        try:
            area = element.get_Parameter(DB.BuiltInParameter.HOST_AREA_COMPUTED).AsDouble() * 0.092903
            length = element.get_Parameter(DB.BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 0.3048
        except:
            pass

        shared_parameter = element_type.LookupParameter("ClassificacaoSecclassSsNumero")
        shared_parameter_value = shared_parameter.AsString()

        shared_parameter2 = element_type.LookupParameter("ClassificacaoSecclassSsDescricao")
        shared_parameter_value2 = shared_parameter2.AsString()

        phase_parameter = element.LookupParameter("Phase Created")
        if phase_parameter is None:
            phase_parameter_value2 = 0
        else:
            phase_parameter_value2 = phase_parameter.AsString()

        Export_BIM_data = {
            "ElementID": element_Id,
            "SECClasS_Code": shared_parameter_value,
            "SECClasS_Title": shared_parameter_value2,
            "Category": category,
            "Family and Type": "{} - {}".format(family_name, type_name),
            "Volume": round(volume, 2),
            "Area": round(area, 2),
            "Length": round(length, 2),
            "Phase_Created": phase_parameter_value2
        }
        data.append(Export_BIM_data)


## SEC-WBS
script_dir = os.path.dirname(__file__) # path do script a correr
parentDirectory = os.path.dirname(script_dir) # path pai do script
parentDirectory2 = os.path.dirname(parentDirectory) # path pai do script
SEC_WBS_file_path = os.path.join(parentDirectory2, 'ByElement.panel', 'FileCreator.pushbutton', 'SEC_WBS.xlsx') # path da pasta e nome do cdv a abrir
SEC_WBS_link = SEC_WBS_file_path
SEC_WBS_table = pd.read_excel(SEC_WBS_link)
SEC_WBS_json = SEC_WBS_table.to_json(orient='records', force_ascii=False).encode('utf8')
SEC_WBS_data = json.loads(SEC_WBS_json)  # Extrat data from JSON

# Tabela excel
Co2Value_path = os.path.join(script_dir, 'Co2Value2.xlsx')
Co2Value_link = Co2Value_path
Co2Value_table = pd.read_excel(Co2Value_link)
Co2Value_json = Co2Value_table.to_json(orient='records', force_ascii=False).encode('utf8')
Co2Value_data = json.loads(Co2Value_json)  # Extrat data from JSON

#2 MAPEAMENTO SECCLASS WBS E MEDIDAS
for i in range(len(data)):
  for row in SEC_WBS_data:
    if data[i]['SECClasS_Code'] is None:
        pass
    else:
        if data[i]['SECClasS_Code'] in row['SECClasS_Code']:
          data[i]['Quantity of elements'] = None
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
          data[i]['Cost'] = 0
          data[i]['GWP_A1-A3'] = row['GWP A1-A3 (Kg/m3, Kg/m2;kg/m;k/U)']


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
data = list_new #passa a ser data


#3 Filtro só classificados
list_classified= []
list_notclassified= []
for row in data:
  if row['SECClasS_Code'] is not None:
    list_classified.append(row)
  else:
    list_notclassified.append(row)

data = list_classified #passa a ser data


## add ao Export BIM coluna do id
for id in range(len(data)):
    data[id]["id"] = id

## Vai buscar o composição por elemento # a passar os elementos medidos à unidade
for i, row in enumerate(data):
    #print(row['Measure'])
    #print(row['id'])
    if row['Measure'] == "U":
        pass
    else:
        element_id = row["ElementID"]
        element = doc.GetElement(DB.ElementId(int(element_id)))
        material_volumes = {}
        material_areas = {}
        for material_id in element.GetMaterialIds(False):
            material = doc.GetElement(material_id)
            material_name = material.Name + " (Volume)"
            material_volume = element.GetMaterialVolume(material_id) * 0.0283168
            material_volumes[material_name] = round(material_volume, 2)

            material_name2 = material.Name + " (Area)"
            material_area = element.GetMaterialArea(material_id, False) * 0.092903
            material_areas[material_name2] = round(material_volume, 2)

        data[i].update(material_volumes)
        data[i].update(material_areas)



########## este é so para ver coisas ############
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
            'GWP_A1-A3': row['GWP_A1-A3'],
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
        #else:
            #print(len(list_V) + len(list_A) + len(list_L) + len(list_U))
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
########################
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



# create an empty list to store material information
materials_list_temp = []

for row in data:
    row['Mass'] = 0
    if row['Measure'] == 'U':
        A = float(row['Conversion_Factor'])
        row['Mass'] = A
        row['Co2_Total'] = row['GWP_A1-A3']
    else:
        # iterate through all elements to extract material information
        materials_set = set()
        for category in categories:
            elements = DB.FilteredElementCollector(doc).OfCategory(category).WhereElementIsNotElementType().ToElements()
            for element in elements:
                for material_id in element.GetMaterialIds(False):
                    material = doc.GetElement(material_id)
                    material_name = material.Name
                    if material_name not in materials_set:
                        material_info = {
                            "Name": material_name + " (Volume)",
                            "Keynote": material.get_Parameter(DB.BuiltInParameter.KEYNOTE_PARAM).AsString()
                        }
                        materials_list_temp.append(material_info)
                        materials_set.add(material_name)

        # criação de lista temporaria e look up para cada material
        for i in range(len(materials_list_temp)):
          for row2 in Co2Value_data:
            if materials_list_temp[i]['Keynote'] is None:
                pass
            else:
                if str(materials_list_temp[i]['Keynote']) == str(row2['SECClasS_Code']):
                    if row2['A1_A3'] is not None:
                        materials_list_temp[i]['Co2'] = float(row2['A1_A3'])
                        materials_list_temp[i]['Conversion_factor'] = row2['Conversion_factor']
                else:
                    pass
        #print(materials_list_temp)
        #print(data)
        for line in materials_list_temp:
            #print(line['Name'])
            if line['Name'] in row:
                string = line['Name']
                nova_string = string[:-9]
                row[nova_string + " (Co2)"] = line['Co2']
                row[nova_string + " (Conversion_factor)"] = line['Conversion_factor']

        # Calculo da massa E CO2)


        #print(materials_list_temp)
        for line in materials_list_temp:
            string = line['Name']
            nova_string = string[:-9]
            if string in row and (nova_string + " (Conversion_factor)") in row:
                row[nova_string + " (Mass)"] = row[string] * row[nova_string + " (Conversion_factor)"]
                row[nova_string + " (M_Co2)"] = row[nova_string + " (Mass)"] * row[nova_string + " (Co2)"]

            else:
                pass

        total_mass = 0
        for key, value in row.items():
            if key.endswith(" (Mass)"):
                total_mass += value
        row["Mass"] = total_mass

        total_co2 = 0
        for key, value in row.items():
            if key.endswith(" (M_Co2)"):
                total_co2 += value
        row["Co2_Total"] = total_co2
###########################
#print(data)
###########  Cascata  PAARA O COST ###########

for i in range(len(data)):
    mass = 0
    row = data[i]
    temp = []

    if 'Measure' in row:
        if row["Conversion_Factor"] is None:
            mass = 0
            cost = 0
        else:
            if row['Measure'] == "V":
                cost = float(str(row['Volume'])) * float(row["Unit_Cost"])

            elif row['Measure'] == "A":
                cost = float(str(row['Area'])) * float(row["Unit_Cost"])

            elif row['Measure'] == "L":
                cost = float(str(row['Length'])) * float(row["Unit_Cost"])

            elif row['Measure'] == "U":
                cost = row["Unit_Cost"]
            else:
                pass
    # Adicionar os valores ao MASS do data
    data[i]['Cost'] = cost
####################
#param_name = "GWP(KgCo2e)"
#for i in range(len(data)):
    #element_id = data[i]['ElementID']
    #element = doc.GetElement(DB.ElementId(int(element_id)))
    #parameter = element.LookupParameter(param_name)
    #print(parameter)
    #t = Transaction(doc, "test_api")
    #t.Start()
    #parameter.Set(round(data[i]["Co2_Total"], 2))
    #print('feito')
    #t.Commit()

########################################
script_dir = os.path.dirname(__file__) # path do script a correr
parentDirectory = os.path.dirname(script_dir) # path pai do script
parentDirectory2 = os.path.dirname(parentDirectory) # path pai do script
CSV_file_path = os.path.join(parentDirectory2, 'ByElement.panel', 'FileCreator.pushbutton', 'Building_Information.csv') # path da pasta e nome do cdv a abrir
## lê o documento da informação do utilizador
column_names2 = ["Project Number",      "Project Name",      "Building Name",      "Building GFA (m2)*",      "Building lifespan (years)*"]

BI = pd.read_csv(CSV_file_path,
                 sep=';',
                 skiprows=1,
                 names=column_names2)

Project_Number = BI.get('Project Number')[0]
Project_Name = BI.get('Project Name')[0]
Building_Name = BI.get('Building Name')[0]
Building_GFA = float(str(BI.get('Building GFA (m2)*')[0]).replace(',','.'))
Building_lifespan = float(str(BI.get('Building lifespan (years)*')[0]).replace(',','.'))




for row in data:
    BLC = float(Building_lifespan) / float(row['Expected_lifespan'])

    row['Normalised requirement factor over building lifetime'] = BLC
    row['BLC_Mass_Total'] = row['Mass'] * BLC
    if row["Co2_Total"] is not None:
        row['BLC_Co2_Total'] = row['Co2_Total'] * BLC
    #print(row['Mass'])

########################################

############# Somatorio de Mass, Co2, Mass_BLC and Co2_BLC
TOTAL_MASS = 0
TOTAL_CO2 = 0
TOTAL_MASS_BLC = 0
TOTAL_CO2_BLC = 0
TOTAL_COST = 0
for row in data:
  TOTAL_MASS = TOTAL_MASS + row['Mass']
  TOTAL_CO2 = TOTAL_CO2 + row['Co2_Total']
  TOTAL_MASS_BLC = TOTAL_MASS_BLC + row['BLC_Mass_Total']
  TOTAL_CO2_BLC = TOTAL_CO2_BLC + row['BLC_Co2_Total']
  SOCIAL_COST = (TOTAL_CO2 / 1000) * 50
  TOTAL_COST = TOTAL_COST + row['Cost']
Normalized_Co2 = (TOTAL_CO2 / (Building_GFA))
Normalized_2 = ( Normalized_Co2 / Building_lifespan)


# Define a separator line
separator = "_______________________________________________________________________________________________________________________________"

# Define a header for the program
header = "........................................................SECCALCULATOR.........................................................."
header_centered = header.center(len(separator))

print(separator)
print(header_centered)
print(separator)
print("Mass Global = {:.2f} kg".format(TOTAL_MASS))
print("GWP Global = {:.2f} kgCo2e".format(TOTAL_CO2))
print("Normalized GWP = {:.2f} kgCo2e/m2".format(Normalized_Co2), end=" ")
print("Social Cost of Carbon = {:.2f} €".format(SOCIAL_COST), end=" ")
print("Mass Global considering Building Life Cycle  = {:.2f} kg".format(TOTAL_MASS_BLC), end=" ")
print("Co2 Global considering Building Life Cycle = {:.2f} kgCo2e".format(TOTAL_CO2_BLC))
print("Budget Estimate = {:.2f} €".format(TOTAL_COST))
print(separator)

print("The CSV and json files were created in the ../ByMaterial.panel/SECCalculator.pushbutton/SECCalculator.pushbutton: "
      "output_data.csv and output_data.json.  ")
print(separator)

        # Adicionar os valores ao MASS do data
   # data[i]['Mass'] = mass
#data[i]['Cost'] = cost
#### all materials names and  ######

###################################
json_file_path = os.path.join(os.path.dirname(__file__), "output_data.json")
csv3_file_path = os.path.join(os.path.dirname(__file__), "output_data.csv")
Output_APP = json.dumps(data)

jsonFile = open(json_file_path, mode='w', newline='')
jsonFile.write(Output_APP)
jsonFile.close()
with open(json_file_path) as json_file:
    data_json = json.load(json_file)

data_csv = open(csv3_file_path, mode='w', newline='', encoding='utf-8')
csv_writer = csv.writer(data_csv)

count = 0
for item in data_json:
    if count == 0:
        header = item.keys()
        csv_writer.writerow(header)
        count += 1

    csv_writer.writerow(item.values())

data_csv.close()