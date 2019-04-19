import pandas as pd
import numpy as np


import sys
print(sys.version)

from pandas import ExcelFile
from pandas import ExcelWriter

import os

os.chdir('/Users/marvinchan/Documents/PythonProgramming')

pd.set_option('display.max_columns', None)

pd.options.display.max_rows = 999

print(os.getcwd())

# show current working directory

print(os.listdir(os.getcwd()))

# show files in current working directory

#input conference room sheets
#input data frames for each sheet

xls = pd.ExcelFile('conferencerooms.xls')

df1 = pd.read_excel(xls, 'Sheet 2')
#df2 = pd.read_excel(xls, 'Week 2')
#df3 = pd.read_excel(xls, 'Week 3')
#df4 = pd.read_excel(xls, 'Week 4')
#df5 = pd.read_excel(xls, 'Week 5')

global bill_hours
bill_hours = []


users = df1['email'].str.split("@", expand= True)
print(users)


users.columns = ['user','domain']

print(users)

companiesunique = users.domain.unique()

print(companiesunique)

def weekly_run():
	for options in companiesunique:
		print(options)
		results = df1[df1['email'].str.contains(options)]
		companyhours = float(sum(results['Hours']))
		bill_hours.append(companyhours)

weekly_run()

print(bill_hours)

companiesdf1 = pd.DataFrame({'Companies': companiesunique, 'Hours' : bill_hours})


print(companiesdf1)


