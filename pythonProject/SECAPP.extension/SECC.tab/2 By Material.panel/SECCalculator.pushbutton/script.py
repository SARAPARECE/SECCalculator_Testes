#! python3
# -*- coding: utf-8 -*-
__title__   = "SECCalulator by material"
__author__  = "Sara Parece"
__doc__     = """This button this button will perform the calculations"""

import clr
import Autodesk.Revit.DB as DB
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.DB import Curve
#import pandas as pd
import json
import csv
import os
from collections import OrderedDict
from Autodesk.Revit.UI import TaskDialog




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
            volume = element.get_Parameter(DB.BuiltInParameter.HOST_VOLUME_COMPUTED).AsDouble() * 0.0283168 # convert to feet to meter m3
        except:
            pass

        try:
            area = element.get_Parameter(DB.BuiltInParameter.HOST_AREA_COMPUTED).AsDouble() * 0.092903  # convert to feet to meter m
            length = element.get_Parameter(DB.BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 0.3048 # convert to feet to meter m
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


####### Tabela csv #### SEC-WBS #######
script_dir = os.path.dirname(__file__) # path do script a correr
parentDirectory = os.path.dirname(script_dir) # path pai do script
parentDirectory1 = os.path.dirname(parentDirectory)
parentDirectory2 = os.path.dirname(parentDirectory1)

SEC_WBS_file_path = os.path.join(parentDirectory2, 'Data', 'SEC_WBS.csv') # path da pasta e nome do cdv a abrir
with open(SEC_WBS_file_path, mode='r', encoding='utf-8-sig') as SEC_WBS:
    SEC_WBS_table = csv.DictReader(SEC_WBS, delimiter=";")
    SEC_WBS_data = list(SEC_WBS_table)
    SEC_WBS_dict = dict(SEC_WBS_data[0])
    list_column_names = list(SEC_WBS_dict.keys())

for i in range(len(SEC_WBS_data)):
    for key in SEC_WBS_data[i]:
        SEC_WBS_data[i][key] = SEC_WBS_data[i][key].replace(',', '.')

########## Tabela excel ########### CO2VALUE
Co2Value_path = os.path.join(parentDirectory2, 'Data', 'Co2Value2.csv')
with open(Co2Value_path, mode='r', encoding='utf-8-sig') as SEC_WBS:
    Co2Value_table = csv.DictReader(SEC_WBS, delimiter=";")
    Co2Value_data = list(Co2Value_table)
for i in range(len(Co2Value_data)):
    for key in Co2Value_data[i]:
        Co2Value_data[i][key] = Co2Value_data[i][key].replace(',', '.')

#ler ficheiro json
path_json = os.path.join(parentDirectory2, 'Data','selected_folder.json')
f = open(path_json)
# returns JSON object as a dictionary
selected_folder1 = json.load(f)
selected_folder = selected_folder1['selected_folder']


#2 MAPEAMENTO SECCLASS WBS E MEDIDAS
for i in range(len(data)):
  for row in SEC_WBS_data:
    if data[i]['SECClasS_Code'] is None:
        pass
    else:
        if data[i]['SECClasS_Code'] == row['SECClasS_Code']:
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
            data[i]['Unit of Measure'] = row['Unit of Measure']
            data[i]['Conversion_Factor'] = row['Conversion Factor (Kg/m3, Kg/m2, kg/m, kg/u)']
            data[i]['Unit_Cost'] = row['Unit_Cost']
            data[i]['Cost'] = 0
            data[i]['GWP_A1-A3'] = row['GWP A1-A3 (kgCo2e/m3, kgCo2e/m2, kgCo2e/m, kgCo2e/u)']


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
    #print(row['Unit of Measure'])
    #print(row['id'])
    if row['Unit of Measure'] == "U":
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
    if row['Unit of Measure'] == 'U':
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
                row[nova_string + " (Mass)"] = float(row[string]) * float(row[nova_string + " (Conversion_factor)"])
                row[nova_string + " (M_Co2)"] = float(row[nova_string + " (Mass)"]) * float(row[nova_string + " (Co2)"])

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
###########  Cascata  PARA O COST ###########

for i in range(len(data)):
    mass = 0
    row = data[i]
    temp = []

    if 'Unit of Measure' in row:
        if row["Conversion_Factor"] is None:
            mass = 0
            cost = 0
        else:
            if row['Unit of Measure'] == "V":
                cost = float(str(row['Volume'])) * float(row["Unit_Cost"])

            elif row['Unit of Measure'] == "A":
                cost = float(str(row['Area'])) * float(row["Unit_Cost"])

            elif row['Unit of Measure'] == "L":
                cost = float(str(row['Length'])) * float(row["Unit_Cost"])

            elif row['Unit of Measure'] == "U":
                cost = row["Unit_Cost"]
            else:
                pass
    # Adicionar os valores ao MASS do data
    data[i]['Cost'] = cost
####################

######################################## ler BI ###
CSV_file_path = os.path.join( selected_folder, 'Building_Information.csv') # path da pasta e nome do cdv a abrir

## lê o documento da informação do utilizador
column_names2 = ["Project Number",      "Project Name",      "Building Name",      "Building GFA (m2)*",      "Building lifespan (years)*"]

with open(CSV_file_path, mode='r', encoding='utf-8-sig') as file:
    reader = csv.DictReader(file, delimiter=';', fieldnames=column_names2)
    BI = list(reader)[1]
try:
    for key in BI:
        BI[key] = BI[key].replace(',', '.')
except Exception as e:
    TaskDialog.Show("ERROR", "Missing information in Building Information.csv.")
    print(f"Error: {str(e)}")



Project_Number = BI['Project Number']
Project_Name = BI['Project Name']
Building_Name = BI['Building Name']
Building_GFA = int(BI['Building GFA (m2)*'])
Building_lifespan = int(BI['Building lifespan (years)*'])




for row in data:
    BLC = float(Building_lifespan) / float(row['Expected_lifespan'])

    row['Normalised requirement factor over building lifetime'] = BLC
    row['BLC_Mass_Total'] = row['Mass'] * BLC
    if row["Co2_Total"] is not None:
        row['BLC_Co2_Total'] = float(row['Co2_Total']) * float(BLC)
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
  TOTAL_CO2 = float(TOTAL_CO2) + float(row['Co2_Total'])
  TOTAL_MASS_BLC = float(TOTAL_MASS_BLC) + float(row['BLC_Mass_Total'])
  TOTAL_CO2_BLC = float(TOTAL_CO2_BLC) + float(row['BLC_Co2_Total'])
  SOCIAL_COST = (TOTAL_CO2 / 1000) * 50
  TOTAL_COST = float(TOTAL_COST) + float(row['Cost'])

Normalized_Co2 = (TOTAL_CO2 / (Building_GFA))
Normalized_2 = ( Normalized_Co2 / Building_lifespan)




###################################
json_file_path = os.path.join(selected_folder, "output_data.json")
csv3_file_path = os.path.join(selected_folder, "output_data.csv")
Output_APP = json.dumps(data)

jsonFile = open(json_file_path, mode='w', newline='')
jsonFile.write(Output_APP)
jsonFile.close()
with open(json_file_path) as json_file:
    data_json = json.load(json_file)

unique_keys = []
for item in data_json:
    for key in item.keys():
        if key not in unique_keys:
            unique_keys.append(key)

with open(csv3_file_path, mode='w', newline='', encoding='utf-8') as data_csv:
    writer = csv.DictWriter(data_csv, fieldnames=unique_keys)
    writer.writeheader()
    for item in data_json:
        ordered_item = OrderedDict()
        for key in unique_keys:
            ordered_item[key] = item.get(key, '')
        writer.writerow(ordered_item)


data_csv.close()

# Iniciando uma transação no Revit
transaction = Transaction(doc, "Atualizar parâmetro GWP(kgCo2e)")
transaction.Start()

for category in categories:
    elements = DB.FilteredElementCollector(doc).OfCategory(category).WhereElementIsNotElementType().ToElements()
    for element in elements:
        element_Id = element.Id.ToString()

        # Procurando o valor de CO2 correspondente na lista de dados
        data_item = next((item for item in data if item["ElementID"] == element_Id), None)

        # Verificando se foi encontrado um valor de CO2 correspondente
        if data_item is not None:
            #print(data_item["Co2_Total"])
            # Verificando se o parâmetro "GWP(kgCo2e)" existe no elemento
            if element.LookupParameter("GWP(kgCo2e)"):
                # Obtendo o parâmetro "GWP(kgCo2e)"
                parameter = element.LookupParameter("GWP(kgCo2e)")
                #print(type(parameter))
                # Atualizando o valor do parâmetro "GWP(kgCo2e)" com o valor de CO2 correspondente
                parameter.Set(data_item["Co2_Total"])

    # Finalizando a transação
transaction.Commit()


########## DIALOG #########
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
import System
from System.Windows.Forms import Application, Form, PictureBox, PictureBoxSizeMode, Label, RichTextBox
from System.Drawing import Image, Icon

# Defina o caminho da imagem
image_path =  os.path.join(parentDirectory2, 'Data', 'SECC.png')

# Conteúdo da caixa de diálogo
content = f"""
Mass Global (A1-A3)   = {TOTAL_MASS:.2f} kg

GWP Global (A1-A3)  = {TOTAL_CO2:.2f} kgCo2e

Normalized GWP (A1-A3)  = {Normalized_Co2:.2f} kgCo2e/m2

Social Cost of Carbon (A1-A3)  = {SOCIAL_COST:.2f} €

Mass Global (A1-A3 + B4)  = {TOTAL_MASS_BLC:.2f} kg

Co2 Global (A1-A3 + B4)  = {TOTAL_CO2_BLC:.2f} kgCo2e

The results are stored in the output.csv and output.json in {selected_folder}.
"""

# Crie uma janela personalizada
class CustomDialog(Form):
    def __init__(self):
        self.InitializeComponent()

    def InitializeComponent(self):
        self.Text = "SECCalculator Results"
        self.Size = System.Drawing.Size(800, 400)

        # Defina o ícone da janela
        Icon_path = os.path.join(parentDirectory2, 'Data', 'SECC (2).ico')
        icon = Icon(Icon_path)
        self.Icon = icon
        # Define a janela para sempre ficar no topo
        self.TopMost = True
        # Crie um PictureBox para exibir a imagem
        picture_box = PictureBox()
        picture_box.Image = Image.FromFile(image_path)
        picture_box.SizeMode = PictureBoxSizeMode.Zoom
        picture_box.Size = System.Drawing.Size(200, 200)
        picture_box.Location = System.Drawing.Point(20, 40)

        # Adicione o PictureBox à janela
        self.Controls.Add(picture_box)

        # Crie um controle RichTextBox para exibir o conteúdo formatado com formatação de texto
        rich_text_box = RichTextBox()
        rich_text_box.Location = System.Drawing.Point(240, 20)
        rich_text_box.Size = System.Drawing.Size(520, 280)
        rich_text_box.ReadOnly = True
        # Adicione o texto ao controle RichTextBox
        rich_text_box.AppendText(content)


        # Adicione negrito às partes desejadas do texto
        self.ApplyBoldText(rich_text_box, "Mass Global (A1-A3)")
        self.ApplyBoldText(rich_text_box, "GWP Global (A1-A3)")
        self.ApplyBoldText(rich_text_box, "Normalized GWP (A1-A3)")
        self.ApplyBoldText(rich_text_box, "Social Cost of Carbon (A1-A3)")
        self.ApplyBoldText(rich_text_box, "Mass Global (A1-A3 + B4)")
        self.ApplyBoldText(rich_text_box, "Co2 Global (A1-A3 + B4)")

        # Adicione o controle RichTextBox à janela
        self.Controls.Add(rich_text_box)

    def ApplyBoldText(self, rich_text_box, target_text):
        index = rich_text_box.Text.find(target_text)
        if index != -1:
            rich_text_box.Select(index, len(target_text))
            rich_text_box.SelectionFont = System.Drawing.Font(rich_text_box.Font, System.Drawing.FontStyle.Bold)


# Crie e mostre a janela personalizada
if __name__ == '__main__':
    dialog = CustomDialog()
    Application.Run(dialog)

##################