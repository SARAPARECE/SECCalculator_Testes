#! python3
# -*- coding: utf-8 -*-

import os
import json
from Autodesk.Revit.UI import TaskDialog

script_dir = os.path.dirname(__file__) # path do script a correr
parentDirectory = os.path.dirname(script_dir) # path pai do script
parentDirectory1 = os.path.dirname(parentDirectory)
parentDirectory2 = os.path.dirname(parentDirectory1)
parentDirectory3 = os.path.dirname(parentDirectory2)
parentDirectory4 = os.path.dirname(parentDirectory3)

#ler ficheiro json

path_json = os.path.join(parentDirectory4, 'Data','selected_folder.json')
f = open(path_json)
# returns JSON object as a dictionary
selected_folder1 = json.load(f)
selected_folder = selected_folder1['selected_folder']
full_path_to_file = os.path.join(selected_folder, 'Building_Information.csv')

try:
    os.startfile(full_path_to_file)
except:
    message = (
        "There is SECClass Codes with out a unit of measure, please confirm in SEC-WBS.csv"
    )
    dialog = TaskDialog("SECCALCULATOR Message")
    dialog.MainInstruction = "Error"
    dialog.MainContent = message
    dialog.Show()