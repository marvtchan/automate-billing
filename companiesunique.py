import pandas as pd
import numpy as np
from pandas import ExcelFile
from pandas import ExcelWriter

pd.options.display.max_rows = 999

xls = pd.ExcelFile('conferencerooms.xls')
df1 = pd.read_excel(xls, 'Sheet 2')

global bill_hours
bill_hours = []

users = df1['email'].str.split("@", expand= True)
users.columns = ['user','domain']
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


