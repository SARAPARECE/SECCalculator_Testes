#! python3
# -*- coding: utf-8 -*-


import clr
import Autodesk.Revit.DB as DB
import sys
#sys.path.append(r'C:\Users\Asus\anaconda3\envs\PyRevit\Lib\site-packages')
#import pandas as pd
import json
import csv
import os.path
from Autodesk.Revit.DB import Transaction
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

for category in categories:
    elements = DB.FilteredElementCollector(doc).OfCategory(category).WhereElementIsNotElementType().ToElements()
    for element in elements:
        element_type = doc.GetElement(element.GetTypeId()) # para propriedades do tipo
        family_name = element_type.FamilyName # para propriedades do elemet o tipo
        element_Id = element.Id.ToString()
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

        shared_parameter3 = element_type.LookupParameter("ClassificacaoSecclassPrNumero")
        if shared_parameter:
            shared_parameter_value3 = shared_parameter.AsString()
        else:
            shared_parameter_value4 = None

        shared_parameter4 = element_type.LookupParameter("ClassificacaoSecclassPrDescricao")
        if shared_parameter2:
            shared_parameter_value4 = shared_parameter2.AsString()
        else:
            shared_parameter_value2 = None

        # SECCLASS CODES - escolher Ss ou Pr
        if shared_parameter3 is None and shared_parameter_value3 == "":
            secclass_code = shared_parameter_value
            secclass_title = shared_parameter_value2
        else:
            secclass_code = shared_parameter_value3
            secclass_title = shared_parameter_value4

        #PHASE CREATED
        phase_parameter = element.LookupParameter("Phase Created")
        if phase_parameter:
            phase_parameter_value2 = None
        else:
            phase_parameter_value2 = None
        #classificacao_parameter = element.get_Parameter("04466190-35a1-4404-82d0-898f32b92ec4")
        #classificacao_value = classificacao_parameter.AsString() if classificacao_parameter else None

        Export_BIM_data = {
            "ElementID": element_Id,
            "SECClasS_Code": secclass_code,
            "SECClasS_Title": secclass_title,
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


###### Tabela csv #### Co2Value ######
Co2Value_file_path = os.path.join(parentDirectory2, 'Data', 'Co2Value2.csv')
with open(Co2Value_file_path, mode='r', encoding='utf-8-sig') as SEC_WBS:
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

#print(Co2Value_data)
#2 MAPEAMENTO SECCLASS WBS E MEDIDAS
for i in range(len(data)):
  for row in SEC_WBS_data:
      for name in list_column_names[2:]:
        if data[i]['SECClasS_Code'] is None:
            pass
        else:
            if data[i]['SECClasS_Code'] == row['SECClasS_Code']:
                data[i][name] = row[name]
                data[i]['Quantity of elements'] = None
                data[i]['Mass'] = None
                data[i]['Co2_Total'] = None
                data[i]['BLC_Mass_Total'] = None
                data[i]['BLC_Co2_Total'] = None
                data[i]['Normalised requirement factor over building lifetime'] = None
                data[i]['Cost'] = None

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



# lê o documento da informação do utilizador
script_dir = os.path.dirname(__file__) # path do script a correr
parentDirectory = os.path.dirname(script_dir) # path pai do script
csv_file_path = os.path.join(selected_folder, 'Building_Elements_Information.csv') # path da pasta e nome do cdv a abrir

row_headers = [
                "ref", "SECClasS_Code", "SECClasS_Title", "Quantity of elements", "Unit of Measure",
                "Conversion Factor (kg/m3, kg/m2, kg/m, kg/u)", "GWP A1-A3 (kgCo2e/m3, kgCo2e/m2, kgCo2e/m, kgCo2e/u)",
                "Unit Cost (€/unit of measure)"
            ]
row_headers.extend(list_column_names[15:])  # Adicione os nomes das colunas restantes
row_headers.append("Total (must be 100%)")


with open(csv_file_path, mode='r', encoding='utf-8-sig') as file:
    reader = csv.DictReader(file, delimiter=';', fieldnames=row_headers)
    EI = list(reader)[1:-1]
for i in range(len(EI)):
    for key in EI[i]:
        EI[i][key] = EI[i][key].replace(',', '.')

########################################

# lê o documento da informação do utilizador

csv_file_path1 = os.path.join(selected_folder, 'Building_Information.csv') # path da pasta e nome do cdv a abrir
column_names2 = ["Project Number",
      "Project Name",
      "Building Name",
      "Building GFA (m2)*",
      "Building lifespan (years)*"]

with open(csv_file_path1, mode='r', encoding='utf-8-sig') as file:
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
####################################

#print("data", data)

###########  Cascata  ###########

for i in range(len(data)):
    mass = 0
    row = data[i]
    temp = []
    for key in range(len(EI)):
        if data[i]["SECClasS_Code"] == EI[key]["SECClasS_Code"]:
            #print(EI[key]["Conversion Factor (kg/m3, kg/m2, kg/m, kg/u)"])
            data[i]["Conversion Factor (kg/m3, kg/m2, kg/m, kg/u)"] = float(EI[key]["Conversion Factor (kg/m3, kg/m2, kg/m, kg/u)"])
            data[i]["Quantity of elements"] = int(EI[key]["Quantity of elements"])
            data[i]["GWP A1-A3 (kgCo2e/m3, kgCo2e/m2, kgCo2e/m, kgCo2e/u)"] = float(EI[key]["GWP A1-A3 (kgCo2e/m3, kgCo2e/m2, kgCo2e/m, kgCo2e/u)"])
            data[i]['Unit of Measure'] = EI[key]["Unit of Measure"]

        else:
            pass

    if 'Unit of Measure' in row:
        if row["Conversion Factor (kg/m3, kg/m2, kg/m, kg/u)"] is None:
            mass = 0
            cost = 0
        else:
            if row['Unit of Measure'] == "V":
                mass = float(str(row['Volume'])) * float(row["Conversion Factor (kg/m3, kg/m2, kg/m, kg/u)"])
                cost = float(str(row['Volume'])) * float(row["Unit_Cost"])

            elif row['Unit of Measure'] == "A":
                mass = float(str(row['Area'])) * float(row["Conversion Factor (kg/m3, kg/m2, kg/m, kg/u)"])
                cost = float(str(row['Area'])) * float(row["Unit_Cost"])

            elif row['Unit of Measure'] == "L":
                mass = float(str(row['Length'])) * float(row["Conversion Factor (kg/m3, kg/m2, kg/m, kg/u)"])
                cost = float(str(row['Length'])) * float(row["Unit_Cost"])

            elif row['Unit of Measure'] == "U":
                mass = row["Conversion Factor (kg/m3, kg/m2, kg/m, kg/u)"]
                cost = row["Unit_Cost"]
            else:
                pass

    # Adicionar os valores ao MASS do data
    data[i]['Mass'] = mass
    data[i]['Cost'] = cost


#print(data[0])
#print(data[0]['Mass'])

#Calculo das Massa por material e add no data
for key in range(len(EI)):
    for lists in data:
      if EI[key]["SECClasS_Code"] == lists['SECClasS_Code']:
        GWP = float(EI[key]['GWP A1-A3 (kgCo2e/m3, kgCo2e/m2, kgCo2e/m, kgCo2e/u)'])
        if GWP == 0 or GWP is None:
            Mass = data[lists['id']]['Mass']
            if Mass is not None:
                for name in list_column_names[15:]:
                    nova_string = name[:-4]
                    data[lists['id']].update({
                        nova_string + ' (Mass)': float(EI[key][name]) * Mass})
                else:
                    pass

# Calculo Co2
for row in data:
    for row2 in Co2Value_data:
      if "Mass" in row and (row['GWP A1-A3 (kgCo2e/m3, kgCo2e/m2, kgCo2e/m, kgCo2e/u)'] is None or float(row['GWP A1-A3 (kgCo2e/m3, kgCo2e/m2, kgCo2e/m, kgCo2e/u)']) == 0):
          for name in list_column_names[15:]:
              nova_string = name[:-4]
              if str(nova_string) == str(row2['SECClasS_Title_EN']):
                  Co2_temp = {
                      nova_string + ' (Co2)': float(row2['A1_A3']) * float(row[nova_string + ' (Mass)'])
                  }
                  data[row['id']].update(Co2_temp)

              else:
                  pass


        # Calculo Co2_GWP
      elif float(row['GWP A1-A3 (kgCo2e/m3, kgCo2e/m2, kgCo2e/m, kgCo2e/u)']) > 0 or not None:
        if row["Unit of Measure"] == "V":
          Co2_GWP = row['Volume'] * row['GWP A1-A3 (kgCo2e/m3, kgCo2e/m2, kgCo2e/m, kgCo2e/u)']
        elif row["Unit of Measure"] == "A":
          Co2_GWP = row['Area'] * float(row['GWP A1-A3 (kgCo2e/m3, kgCo2e/m2, kgCo2e/m, kgCo2e/u)'])
        elif row["Unit of Measure"] == "L":
          Co2_GWP = row['Length'] * float(row['GWP A1-A3 (kgCo2e/m3, kgCo2e/m2, kgCo2e/m, kgCo2e/u)'])
        elif row["Unit of Measure"] == "U":
          Co2_GWP = float(row['GWP A1-A3 (kgCo2e/m3, kgCo2e/m2, kgCo2e/m, kgCo2e/u)'])
        else:
          Co2_GWP = 0
        data[row['id']]['Co2_Total'] = Co2_GWP
        #print(row['id'], row['SECClasS_Code'], "Co2_GWP=", Co2_GWP)

    total_co2 = 0
    for key, value in row.items():
        if key.endswith(" (Co2)"):
            total_co2 += value
    row["Co2_Total"] = total_co2

# LP * Expected_lifespan e add vari LP_factor * mass
for row in data:
    for name in list_column_names[15:]:
        nova_string = name[:-4]
        BLC = int(Building_lifespan) / int(row['Expected_lifespan'])
        data[row['id']]['Normalised requirement factor over building lifetime'] = BLC
        data[row['id']]['BLC_Mass_Total'] = float(row['Mass']) * float(BLC)
        if row["Co2_Total"] is not None:
            data[row['id']]['BLC_Co2_Total'] = row['Co2_Total'] * BLC
            if nova_string + " (Mass)" in row:
                data[row['id']].update({
                  nova_string +" (Mass BLC)": BLC * row[nova_string + " (Mass)"]
                  })
            if nova_string + " (Co2)" in row:
                data[row['id']].update({
                        nova_string + " (Co2 BLC)": BLC * row[nova_string + " (Co2)"]
                        })
###############################################################################
#print(data)
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
  TOTAL_COST = float(TOTAL_COST) + float(row['Cost'])
  SOCIAL_COST = float((TOTAL_CO2 / 1000)) * 50
Normalized_Co2 = float(TOTAL_CO2) / float(Building_GFA)
Normalized_2 = (Normalized_Co2 / Building_lifespan)
#########



################## Export data ################
for i in range(len(data)):
    for key in data[i]:
        data[i][key] = str(data[i][key]).replace('.', ',')

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

# Iniciando uma transação no Revit
transaction = Transaction(doc,"Atualizar parâmetro GWP(kgCo2e)")
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
image_path = os.path.join(parentDirectory2, 'Data', 'SECC.png')

# Conteúdo da caixa de diálogo
content = f"""
Mass Global   = {TOTAL_MASS:.2f} kg

GWP Global = {TOTAL_CO2:.2f} kgCo2e

Normalized GWP = {Normalized_Co2:.2f} kgCo2e/m2

Social Cost of Carbon = {SOCIAL_COST:.2f} €

Total Cost  = {TOTAL_COST:.2f} €

Mass Global considering Building Life Cycle  = {TOTAL_MASS_BLC:.2f} kg

Co2 Global considering Building Life Cycle  = {TOTAL_CO2_BLC:.2f} kgCo2e

The results are stored in the output.csv and output.json in {selected_folder} .
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
        icon = Icon(Icon_path)  # Substitua pelo caminho do ícone desejado
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
        self.ApplyBoldText(rich_text_box, "Mass Global")
        self.ApplyBoldText(rich_text_box, "GWP Global")
        self.ApplyBoldText(rich_text_box, "Normalized GWP")
        self.ApplyBoldText(rich_text_box, "Social Cost of Carbon")
        self.ApplyBoldText(rich_text_box, "Total Cost")
        self.ApplyBoldText(rich_text_box, "Mass Global considering Building Life Cycle")
        self.ApplyBoldText(rich_text_box, "Co2 Global considering Building Life Cycle")

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