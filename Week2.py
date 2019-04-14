import pandas as pd
import numpy as np
import random 
import itertools

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


#create variable for Mammoth

options1 = ['AesculaTech']

options2 = ['Alaunus']

options3 = ['Alpine Roads','Alpine Roads, Inc.']

options4 = ['Applaud Medical']

options5 = ["Coral Genomics"]

options6 = ["Cypre, Inc."]

options7 = ["Deciduous Therapeutics"]

options8 = ["Delve Therapeutics"]

options9 = ['Epiodyne']

options10 = ["Fountain Therapeutics"]

options11 = ['GeneTether, Inc']

options12 = ['Gordian Biotechnology']

options13 = ['GraphWear Technologies Inc.']

options14 = ["Logic.Ink", "LogicInk"]

options15 = ["Mammoth Diagnostics", "Mammoth BioSciences"]

options16 = ['Mitokinin INC']

options17 = ['Mojo Health Inc']

options18 = ['Naked Biome']

options19 = ['Nitrome Biosciences']

options20 = ['OneSkin']

options21 = ['Prellis Biologics']

options22 = ['Provenance Biofabrics, Inc','Provenance','Provenance Biofabrics Inc.']

options23 = ["Quartz Therapeutics"]

options24 = ['Rumi Scientific']

options25 = ['Scaled Biolabs','Scaled Biolabs Inc.']

options26 = ["Scribe Biosciences"]

options27 = ['Siolta','Siolta Therapeutics']

options28 = ["Soteria Biotherapeutics",'Soteria Biotherapeutics, Inc.']

options29 = ["SyntheX, Inc."]

options30 = ['Telo Therapeutics, Inc.']

options31 = ['The Wild Type']


companies = [options1, options2, options3, options4, options5, options6, options7, options8, options9, 
options10, options11, options12, options13, options14, options15, options16, options17,
options18, options19, options20, options21, options22, options23, options24, options25, options26,
options27, options28, options29, options30, options31]




print(companies)




def weekly_run():
	for options in companies:
		print(options)
		results = df1.loc[df1['Organization'].isin(options)]
		companyhours = float(sum(results['Hours']))
		bill_hours.append(companyhours)





weekly_run()




		
print(bill_hours)



		#bill_hours.append(total_hours)


		







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
