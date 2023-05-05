#! python3
# -*- coding: utf-8 -*-

# Define a separator line
separator = "_______________________________________________________________________________________________________________________________"

# Define a header for the program
header = "........................................................SECCALCULATOR.........................................................."
header_centered = header.center(len(separator))

instructions1 = """
1st - The model has no materials or material layers and the building elements are identified with a single generic element.
2nd - All materials and products are modelled with independent elements. 
3rd - Building elements are modelled with compound elements, with layers of materials.
"""

instructions2 = """
A. Identify the modelling method used, if 1st or 2nd: 
B. Verify that all model elements are classified
1. Open Building_Information.csv and Building_Elements_Information.csv files in Excel.
2. Fill in the required information Building_Information.csv and save.
3. If the information is not previously in the SEC_WBS database, you will need to fill 
in Building_Elements_Information.csv the following queries: the material composition of 
each element in % each element or the global warming potential (kgCo2eq/unit of measurement) 
and the conversion factor (kg/unit of measurement). You can populate the database - SEC_WBS.xlsx - 
with this information for each SECClass code. 
If the database has the information for that code, this step is not necessary anymore; just a check. 
This database is an Excel document available in the FileCreator.pushbutton folder and can be changed 
whenever needed. 
Suppose for the same code; there are several different construction elements. 
In that case, these can be added to the SEC_WBS database by adding a suffix for example, Ss_00_00_00.PDE.
4. If all data is filled, click the SECCalculator button
5. Then the results with the totals appear in the dashboard and the results for all  elements are 
saved in the documents output_da.csv and output.json.

"""
instructions3 = """
A. Identify the modeling method used, if 3rd: 
B.  Check that all model elements are classified
1. Open the Building_Information.csv files in Excel.
2. Fill in the required Building_Information.csv and save.
3. Verify that the Co2Value document has all the correct data
4. If all the data is filled, click on the SECCalculator - by material button
5. Then the results with the totals appear in the dashboard and the results for 
all the model elements are saved in the documents output_data.csv and output.json. 
These files are automatically saved created in the FileCreator.pushbutton. 
"""

print(separator)
print(header_centered)
print(separator)
print("There are 3 modelling methods and, depending on each, different approaches to the calculation must be considered:")
print(instructions1)
print(separator)
print("Instructions:")
print(instructions2)
print(separator)
print(instructions3)
print(separator)
print("This was bring to you by Sara Parece. For more information: www.secclass.pt")