import pandas as pd
import numpy as np
import random 
import sys


import os
print(os.getcwd())

# show current working directory

print(os.listdir(os.getcwd()))

# show files in current working directory

#input conference room sheets
#input data frames for each sheet

xls = pd.ExcelFile('conferencerooms.xlsx')
df1 = pd.read_excel(xls, 'Week 5')
df2 = pd.read_excel(xls, 'Week 2')
df3 = pd.read_excel(xls, 'Week 3')
df4 = pd.read_excel(xls, 'Week 4')
#df5 = pd.read_excel(xls, 'Week 5')



all_df = []
for key in df.keys():
    all_df.append(df[key])