import os 
import pandas as pd
import warnings
import sys
from datetime import datetime
from time import localtime, strftime

start_time = datetime.now()

#Warning Suppression

warnings.filterwarnings("ignore", message="Workbook contains no default style")
warnings.filterwarnings("ignore", message="The behaviour of Dataframe concatenation with empty or all-NA entries is deprecated. In a future version, this will no longer exclude empty or all-NA columns when determining the result dtypes. To retain old behaviour, exclude the relevant entries before the concat operation")


# Routes of input folder, if one does not exist it will make one.
file_route = r'' # File location of folder labelled "input" which contains the excel files

input_exist_bool = os.path.exists(file_route+'\\input')
if input_exist_bool != True:
    os.makedirs(file_route+'\\input')
else:
    pass

#FIle Path

file_path = (file_route + r"/input")
first_file_path = (os.listdir(file_path)[0])
overall_file_path = file_path + "\\" + first_file_path

#Grabbing headers from first file

excel_file = pd.ExcelFile(overall_file_path) #Path to first Excel file
sheet_names = excel_file.sheet_names # Grabs worksheet names.

#Logs folder
logs_exist_bool = os.path.exists(file_route+"\\logs")
if logs_exist_bool != True:
    os.makedirs(file_route + "\\logs")
else:
    pass

#Terminal Logging
actual_time = strftime("%Y-%m-%d %H-%M-%S",localtime())
f = open(file_route + r'\logs' + '\\' + str(actual_time) + ".txt",'w')
sys.stdout = f

#Iteration through all files 

df = []
for file in os.listdir(file_path):
    if file.endswith('.xlsx'):
        print ('Loading file {0}...'.format(file))
        for item in sheet_names:
            read = pd.read_excel(os.path.join(file_path,file),sheet_name = item)
            df.append(read)
            print(item,"has",read.shape[0],"many rows")

file_name = file_route + r'\Merged_file.csv'
print("Creating file",file_name,'...')
df_master = pd.concat(df, axis = 0)
df_master.to_csv(file_name)
print("Files have been merged")

size = df_master.shape[0]
print('The dataframe contains',size,'rows')
end_time = datetime.now()
print("Terminal Logged \n Code finished with runtime",format(end_time - start_time) + 'hr/min/sec')
f.close()