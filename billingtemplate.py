import pandas as pd
import numpy as np
import random 

from pandas import ExcelFile
from pandas import ExcelWriter

import os

pd.set_option('display.max_columns', None)
print(os.getcwd())

# show current working directory

print(os.listdir(os.getcwd()))

# show files in current working directory

#input conference room sheets
#input data frames for each sheet

xls = pd.ExcelFile('conferencerooms.xlsx')
df1 = pd.read_excel(xls, 'Week 2')
#df2 = pd.read_excel(xls, 'Week 2')
#df3 = pd.read_excel(xls, 'Week 3')
#df4 = pd.read_excel(xls, 'Week 4')
#df5 = pd.read_excel(xls, 'Week 5')

global bill_hours
bill_hours = []

companies = [['AesculaTech'], ['Alaunus'], ['Alpine Roads','Alpine Roads, Inc.'], ['Applaud Medical'], ["Coral Genomics"], ["Cypre, Inc."], ["Deciduous Therapeutics"], ["Delve Therapeutics"], ['Epiodyne'], 
["Fountain Therapeutics"], ['GeneTether, Inc'], ['Gordian Biotechnology'], ['GraphWear Technologies Inc.'], ["Logic.Ink", "LogicInk"], ["Mammoth Diagnostics", "Mammoth BioSciences", 'Mammoth Biosciences'], ['Mitokinin INC'], ['Mojo Health Inc'],
['Naked Biome'], ['Nitrome Biosciences'], ['OneSkin'], ['Prellis Biologics'], ['Provenance Biofabrics, Inc','Provenance','Provenance Biofabrics Inc.'], ["Quartz Therapeutics"], ['Rumi Scientific']
, ['Scaled Biolabs','Scaled Biolabs Inc.'], ["Scribe Biosciences"],
['Siolta','Siolta Therapeutics'], ["Soteria Biotherapeutics",'Soteria Biotherapeutics, Inc.'], ["SyntheX, Inc."], ['Telo Therapeutics, Inc.'], ['The Wild Type']]

def weekly_run():
	for options in companies:
		print(options)
		results = df1.loc[df1['Organization'].isin(options)]
		companyhours = float(sum(results['Hours']))
		bill_hours.append(companyhours)


weekly_run()

print(bill_hours)

check1 = sum(bill_hours)

print(check1)


companiesdf1 = pd.DataFrame({'Companies': ['AesculaTech','Alaunus','Alpine Roads','Applaud',
	'Coral Genomics','Cypre','Deciduous','Delve','Epiodyne','Fountain','GeneTether','Gordian','GraphWear',
	'LogicInk','Mammoth','Mitokinin','Mojo','Naked Biome','Nitrome','OneSkin','Prellis','Provenance',
	'Quartz','Rumi','Scaled BioLabs','Scribe','Siolta','Soteria','SyntheX','Telo','The Wild Type'], 'Hours' : bill_hours})


companiesdf1.to_excel('output2.xlsx')

output2 = pd.read_excel('output2.xlsx', index_col=None, 
	na_values=['NA'], usecols = 'B:C')

print (output2)

m1 = output2.iloc[:,1].ge(4)
m2 = output2.iloc[:,1].le(4)

print(output2)

output2.loc[m1,'Hours'] -= 4
output2.loc[m2,'Hours'] = 0


print('\nResult Week2 Billable :\n', output2)

output2.to_excel('week2bill.xlsx')
