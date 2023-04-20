##########Prequisites##########
import os

try:
    import win32com.client
except ImportError:
    import pip
    pip.main(['install', 'pywin32'])
finally:
    import win32com.client

try:
    import pandas as pd
except ImportError:
    import pip
    pip.main(['install', 'pandas'])
finally:
    import pandas as pd

try:
    import openpyxl
except ImportError:
    import pip
    pip.main(['install', 'openpyxl'])
finally:
    import openpyxl


##########Photoshop Dispatch Object##########
psApp=win32com.client.Dispatch("Photoshop.Application")

##########Welcome Message##########
print("**********CERTIFICATE MAKER**********")
print("Welcome to CrtMaker 1.01")
print("This script automatically issues certificates from a PSD template file with names extracted from a MS Excel file")
print("Created by Soroush Dianaty (soroushdianaty@gmail.com)")
print("Free and open-source under CC BY-NC 2.0 licence")
print("*************************************\n")

##########Load PSD file##########      
# Specify PSD path
psd_path = input('Please enter PSD file path or press 1 for the current path: ')
if psd_path == '1':
    psd_path = os.getcwd()  # Set path to current if 1 is entered

# Select PSD file
psd_files = []
for file in os.listdir(rf'{psd_path}'):
    if file.endswith(".psd"):
        psd_files.append(file)
print('\n*****AVAILABLE FILES*****')
for index,value in enumerate(psd_files):
    print(f'{index}: {value}')
print('*************************\n')
        
select_template = int(input('Please choose your template: '))

# Create COM object and load layers
psApp.Open(fr"{psd_path}\{psd_files[select_template]}")
doc = psApp.Application.ActiveDocument

# Select layer
layers = []
for layer in doc.layers:
    layers.append(layer.name)
print('\n*****AVAILABLE LAYERS*****')
for index,value in enumerate(layers):
    print(f'{index}: {value}')
print('**************************\n')
    
layer_name = int(input('Please select a layer: '))

layer = doc.ArtLayers[layers[layer_name]]
text_of_layer = layer.TextItem


##########Load names##########
#Select Excel
excel_path = input('Please choose the Excel file path or press 1 for the current path: ')
if excel_path == '1':
    excel_path = os.getcwd()  # Set path to current if 1 is entered
    
excel_files = []
for file in os.listdir(rf'{excel_path}'):
    if file.endswith(".xls") or file.endswith(".xlsx"):
        excel_files.append(file)
print('\n*****AVAILABLE FILES*****')
for index,value in enumerate(excel_files):
    print(f'{index}: {value}')
print('*************************\n')
        
select_excel = int(input('Please choose your Excel: '))

# Select sheet
if len(pd.ExcelFile(rf'{excel_path}\{excel_files[select_excel]}').sheet_names) > 1:
    sheets = []
    for sheet in pd.ExcelFile(rf'{excel_path}\{excel_files[select_excel]}').sheet_names:
        sheets.append(sheet)
    print('\n*****AVAILABLE SHEETS*****')
    for index,value in enumerate(sheets):
        print(f'{index}: {value}')
    print('**************************\n')
    select_sheet = int(input('Please select a sheet: '))
else:
    select_sheet = 0

# Select name range
column = input('Select column (A-Z, Capital letter): ')
first_row = int(input('Select first row:  '))-1
last_row = int(input('Select last row:  '))
column_dict = {'A':0,'B':1, 'C':2, 'D':3, 'E':4, 'F':5, 'G':6, 'H':7,
               'I':8, 'J':9, 'K':10, 'L':11, 'M':12, 'N':13, 'O':14,
               'P':15, 'Q':16, 'R':17, 'S':18, 'T':19, 'U': 20, 'V':21,
               'W':22, 'X':23, 'Y':24, 'Z':25}
excel = pd.read_excel(rf'{excel_path}\{excel_files[select_excel]}',select_sheet, header=None)
names = excel.iloc[first_row:last_row, column_dict[column]]
names = [name for name in names if type(name)==str]
print('Selected names:', '\n')
for index,value in enumerate(names):
    print(f'{index}: {value}')
print('\n')
##########PNG specifications##########
options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
options.Format = 13   # PNG Format
options.PNG8 = False  # Sets it to PNG-24 bit

##########Output path##########
# Select output path
output_path = input("Please choose the output path or press 1 for the current path: ")
print('\n')
if output_path == '1':
    output_path = os.getcwd()  # Set path to current if 1 is entered

# Create output path
if not os.path.exists(rf"{output_path}\Certificates"):
    path = rf"{output_path}\Certificates"
else:
    for i in range(2,20):
        if not os.path.exists(rf"{output_path}\Certificates {i}"):
            path = rf"{output_path}\Certificates {i}"
            break
    
os.makedirs(path)

##########LOOP##########
for name in names:
    text_of_layer.contents = name  # Issue certificate
    doc.Export(ExportIn=fr"{path}\{name}.png", ExportAs=2, Options=options)  # Export PNG
    print(f'Certificate for {name} issued successfully!')

##########FINISHED##########
print('All Done!')

