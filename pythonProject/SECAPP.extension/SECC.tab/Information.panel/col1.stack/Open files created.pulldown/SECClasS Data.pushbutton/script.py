#! python3
# -*- coding: utf-8 -*-

import os

script_dir = os.path.dirname(__file__) # path do script a correr
parentDirectory = os.path.dirname(script_dir) # path pai do script
parentDirectory1 = os.path.dirname(parentDirectory)
parentDirectory2 = os.path.dirname(parentDirectory1)
parentDirectory3 = os.path.dirname(parentDirectory2)
parentDirectory4 = os.path.dirname(parentDirectory3)
full_path_to_file = os.path.join(parentDirectory4, 'Data', 'SEC_WBS.csv')
os.system(full_path_to_file)